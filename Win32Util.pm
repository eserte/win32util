# -*- perl -*-

#
# $Id: Win32Util.pm,v 1.18 2001/01/24 21:25:17 eserte Exp $
# Author: Slaven Rezic
#
# Copyright (C) 1999, 2000, 2001 Slaven Rezic. All rights reserved.
# This package is free software; you can redistribute it and/or
# modify it under the same terms as Perl itself.
#et
# Mail: eserte@cs.tu-beetrlin.de
# WWW:  http://user.cs.tu-berlin.de/~eserte/
#

package Win32Util;

=head1 NAME

Win32Util - a collection of Win32 related functions

=head1 SYNOPSIS

    use Win32Util;

=head1 DESCRIPTION

This is a collection of Win32 related functions. There are no strict
prerequirements for this module, however, full functionality can only
be achieved if some CPAN modules (Win32::Registry, Win32::API,
Win32::DDE, Win32::Shortcut ...) are available. By default, most of
these modules are already bundled with the popular ActivePerl package.

=cut

use strict;
use vars qw($DEBUG $browser_ole_obj $VERSION);

$VERSION = sprintf("%d.%02d", q$Revision: 1.18 $ =~ /(\d+)\.(\d+)/);
$DEBUG=0 unless defined $DEBUG;

# XXX Win-Registry-Funktionen mit Hilfe von Win32::API und
# der Hilfe von der Access-Webpage nachbilden...

# Laut Microsoft-Dokumentation soll für den Ort des Programm-Verzeichnisses
# die Funktion
#     SHGetSpecialFolderLocation(..., CSIDL_PROGRAMS, ...)
# verwendet werden.

use vars qw(%API_FUNC %API_DEF);

%API_DEF = ("SystemParametersInfo" => {Lib => "user32",
				       In  => ['N', 'N', 'P', 'N'],
				       Out => 'N'},
	    "GetSystemMetrics"     => {Lib => "user32",
				       In  => ['N'],
				       Out => 'N'},
	    "SHAddToRecentDocs"    => {Lib => "shell32",
				       In  => ['I', 'P'],
				       Out => 'I'},
	    "GetLogicalDrives"     => {Lib => "kernel32",
				       In  => [],
				       Out => "I"},
	    "GetDriveType"         => {Lib => "kernel32",
				       In  => ["P"],
				       Out => "I"},
	    "GetSysColor"          => {Lib => "user32",
				       In  => ['N'],
				       Out => 'N'},
	    "GetUserName"          => {Lib => "advapi32",
				       In  => ['P', 'P'],
				       Out => 'I'},
	   );

sub _get_api_function {
    my $name = shift;
    eval {
	require Win32::API;
	if (!exists $API_FUNC{$name}) {
	    my $def = $API_DEF{$name};
	    if (!$def) {
		die "No API definition for $name";
	    }
	    $API_FUNC{$name} = new Win32::API ($def->{Lib}, $name,
					       $def->{In}, $def->{Out});
	}
    };
    warn $@ if $@;
    $API_FUNC{$name};
}

=head1 PROGRAM EXECUTION FUNCTIONS

=head2 start_any_viewer($file)

Based on extension of the given $file, start the appropriate viewer.

=cut

sub start_any_viewer {
    my $file = shift;
    require File::Basename;
    my($n,$p,$suffix) = File::Basename::fileparse($file, "\.[^.^]*");
    if ($suffix =~ /^html?$/) {
	return start_html_viewer($file);
    } elsif ($suffix eq 'ps') {
	return start_ps_viewer($file);
    } else {
	my $class = get_class_by_ext($suffix);
	if ($class) {
	    my $cmd = get_reg_cmd($class);
	    if (!$cmd) {
		warn "No command for class $class";
	    } else {
		return start_cmd($cmd, $file);
	    }
	} else {
	    warn "Can't start viewer for $file";
	}
    }
    0;
}

=head2 start_html_viewer($file)

Start a html viewer with the given file. This is mostly a WWW browser.

=cut

sub start_html_viewer {
    my $file = shift;
    if (!start_html_viewer_cmd($file)) {
        if (!start_html_viewer_dde($file)) {
            system("netscape $file &");
            if ($?) { return undef }
        }
    }
    1;
}

sub start_html_viewer_cmd {
    my $file = shift;
    my $html_viewer = get_html_viewer();
    if ($html_viewer =~ /netscape/i) {
	# Bei Netscape: HTML-Viewer funktioniert auf Dateien, nicht auf URLs!
	$file =~ s/^file://;
    }
    start_cmd($html_viewer, $file);
}

=head2 start_ps_viewer($file)

Start a postscript viewer with the given file.

=cut

sub start_ps_viewer {
    my $file = shift;
    if (!start_ps_viewer_cmd($file)) {
        system("gsview32 $file %");
        if ($?) { return undef }
    }
    1;
}

sub start_ps_viewer_cmd {
    my $file = shift;
    my $ps_viewer = get_ps_viewer();
    start_cmd($ps_viewer, $file);
}

=head2 start_ps_print($file)

Print a postscript file via a postscript viewer.

=cut

sub start_ps_print {
    my $file = shift;
    if (!start_ps_print_cmd($file)) {
        system("gsview32 /p $file %");
        if ($?) { return undef }
    }
    1;
}

sub start_ps_print_cmd {
    my $file = shift;
    my $ps_print = get_ps_print();
    start_cmd($ps_print, $file);
}

sub start_html_viewer_dde {
    my $file = shift;
    my ($app, $topic) = get_html_viewer_dde();
    start_dde($app, $topic, $file);
}

# XXX change to use IE or NS
# XXX Test it...
# Return a Win32::OLE object. With the XXX
sub show_browser_file {
    require Win32::OLE ;
    my $file = shift;
    if (!defined $browser_ole_obj) {
	$browser_ole_obj = Win32::OLE->new('InternetExplorer.Application');
    }
    if (defined $file && defined $browser_ole_obj) {
	$browser_ole_obj->Navigate($file);
    }
}

=head2 start_mail_composer($mailaddr)

Start a mail composer with $mailaddr as the recipient.

=cut

sub start_mail_composer {
    my $mailaddr = shift;
    my $mailto_cmd = get_mail_composer();
    start_cmd($mailto_cmd, $mailaddr);
}

sub get_html_viewer {
    my $class = get_class_by_ext(".htm") || "htmlfile";
    get_reg_cmd($class);
}
sub get_ps_viewer {
    my $class = get_class_by_ext(".ps") || "psfile";
    get_reg_cmd($class);
}
sub get_ps_print {
    my $class = get_class_by_ext(".ps") || "psfile";
    get_reg_cmd($class, "print");
}
sub get_mail_composer {
    my $cmd = get_reg_cmd("mailto");
    if ($cmd) {
	return $cmd;
    } else {
	eval <<'EOF';
	   use Win32::Registry;
	   my($key_ref, $key_ref2);
	   my $root = "SOFTWARE\\Clients\\Mail";
	   return unless $main::HKEY_LOCAL_MACHINE->Open($root, $key_ref);
	   my $key_ref2 = [];
	   return unless $key_ref->GetKeys($key_ref2);
	   my $clients = [@$key_ref2];
	   my $hashref;
	   if ($key_ref->GetValues($hashref)) {
	   	unshift @$clients, $hashref->{""}[2]; # default mailer
	   }
	   foreach my $client (@$clients) {
	       if ($main::HKEY_LOCAL_MACHINE->Open("$root\\$client\\Protocols\\mailto\\shell\\open\\command", $key_ref)) {
	       	   my $hashref;
	           if ($key_ref->GetValues($hashref)) {
		       $cmd = $hashref->{""}[2];
		       last;
	           }
	       }
	   }
EOF
	warn $@ if $@;
	return $cmd if defined $cmd;
	die "Can't send mail";
    }
}

sub get_html_viewer_dde {
    eval q{
        use Win32::Registry;
        my($app_ref, $topic_ref);
        return unless $main::HKEY_CLASSES_ROOT->Open('htmlfile\shell\open\ddeexec\Application', $app_ref);
        return unless $main::HKEY_CLASSES_ROOT->Open('htmlfile\shell\open\ddeexec\Topic', $topic_ref);
        my($app_hashref, $topic_hashref);
        return unless $app_ref->GetValues($app_hashref);
        return unless $topic_ref->GetValues($topic_hashref);
        ($app_hashref->{""}[2], $topic_hashref->{""}[2]);
    };
}

=head2 start_cmd($cmd, @args...)

Start an external program named $cmd. $cmd should be the full path to
the executable. @args are passed to the program. The program is
spawned, that is, executed in the background.

=cut

sub start_cmd {
    my($fullcmd, @args) = @_;

    my($appname, $base, $cmdline);
    eval q{
        use File::Basename;
        use Text::ParseWords;
        my(@words) = parse_line('\s+', 1, $fullcmd);
        $appname = shift @words; $appname =~ s/\"//g;
        my $argstr = join(" ", @words);
        $base = basename($appname); $base =~ s/\"//g;
        $cmdline = $base;
        my %arg_used;
        $argstr =~ s/(%(\d))/ $arg_used{$2-1}=1; defined($args[$2-1]) ? $args[$2-1] : "" /eg;
        $cmdline .= " $argstr";
        for my $i (0 .. $#args) {
            if (!$arg_used{$i}) {
                $cmdline .= " $args[$i]";
            }
        }
        warn "start_cmd: " . $cmdline . "\n" if $DEBUG;
    };
    warn $@ if $@;

    my $r;
    eval q{
        use Win32::Process;
        my $proc;
        $r = Win32::Process::Create($proc, $appname, $cmdline,
				    0, NORMAL_PRIORITY_CLASS, ".");
    };
    if ($@) { # try Win32::Spawn (built-in)
        my $pid;
        $r = Win32::Spawn($appname, $cmdline, $pid);
    }
    $r;
}

=head2 start_dde($app, $topic, $arg)

Start a program via DDE. (What is $app and $topic?)

=cut

sub start_dde {
    my($app, $topic, $arg) = @_;
    my $r;
    eval q{
        use Win32::DDE::Client;
# XXX
$app="Netscape";# geht nur mit Netscape und nicht mit "Netscape 4.0" - warum?
        my $dde = new Win32::DDE::Client($app, $topic);
        warn "DDE Client with $app and $topic: $dde\n" if $DEBUG;
        if ($dde->Error) {
            warn "Unable to initiate DDE connection: " . $dde->Error . "\n";
            return;
        }
        $r = $dde->Request($arg);
        $dde->Disconnect;
    };
    $r;
}

=head1 EXTENSION AND MIME FUNCTIONS

=head2 get_reg_cmd($filetype[, $opentype])

Get a command from registry for $filetype. The "open" type is
returned, except stated otherwise.

=cut

sub get_reg_cmd {
    my($filetype, $opentype) = @_;
    $opentype = 'open' if !defined $opentype;
    my $cmd;
    eval q{
        use Win32::Registry;
        my($reg_key, $key_ref, $hashref);
        $reg_key = join('\\\\', $filetype, 'shell', $opentype, 'command');
        return unless $main::HKEY_CLASSES_ROOT->Open($reg_key, $key_ref);
        return unless $key_ref->GetValues($hashref);
        $cmd = $hashref->{""}[2];
    };
    warn $@ if $@;
    $cmd;
}

=head2 get_class_by_ext($ext)

Return the class name for the given extension.

=cut

sub get_class_by_ext {
    my $ext = shift;
    my $class;
    eval q{
        use Win32::Registry;
        my($key_ref, $hashref);
        return unless $main::HKEY_CLASSES_ROOT->Open($ext, $key_ref);
        return unless $key_ref->GetValues($hashref);
        $class = $hashref->{""}[2];
    };
    warn $@ if $@;
    $class;
}

=head2 install_extension(%args)

Install a new extension (class) to the registry. The function may take
the following key-value parameters:

=over 4

=item -extension

Required. The extension to be installed. The extension should start
with a dot. This can also be an array reference to a number of
extensions.

=item -name

Required. The class name of the new extension. May be something like
Excel.Application.

=item -icon

The (full) path to a default icon file (format should be .ico).

=item -open

The default open command (used if the file is double-clicked in the
explorer).

=item -print

The default print command.

=item -desc

An optional description.

=item -mime

The mime type of the extension (something like text/plain).

=back

=cut

sub install_extension {
    my(%args) = @_;
    my $ext  = $args{-extension} or die "Missing -extension parameter";
    my @ext;
    push @ext, (ref $ext eq 'ARRAY' ? @$ext : $ext);
    foreach my $ext (@ext) {
        if ($ext !~ /^\./) {
	    warn "Extension $ext does not start with dot";
	}
    }
    my $name  = $args{-name} or die "Missing -name parameter";
    my $icon  = $args{-icon};
    my $open  = $args{"-open"};
    my $print = $args{"-print"};
    my $desc  = $args{"-desc"};
    my $mime  = $args{"-mime"};
    eval q{
	use Win32::Registry;
	foreach my $ext (@ext) {
	    my $ext_reg;
	    $main::HKEY_CLASSES_ROOT->Create($ext, $ext_reg);
	    $ext_reg->SetValue("", REG_SZ, $name);
	    if (defined $mime) {
		$ext_reg->SetValueEx("Content Type", 0, REG_SZ, $mime);
	    }
	}
	my $name_reg;
	$main::HKEY_CLASSES_ROOT->Create($name, $name_reg);
	if (defined $desc) {
	    $name_reg->SetValue("", REG_SZ, $desc);
	}
	if (defined $icon) {
	    my $icon_reg;
	    $name_reg->Create("DefaultIcon", $icon_reg);
	    $icon_reg->SetValue("", REG_SZ, $icon);
	}
	my $shell_reg;
	if (defined $open && defined $print) {
	    $name_reg->Create("shell", $shell_reg);
	}
	if (defined $open) {
	    my $open_reg;
	    $shell_reg->Create("open", $open_reg);
	    my $command_reg;
	    $open_reg->Create("command", $command_reg);
	    $command_reg->SetValue("", REG_SZ, $open);
	}
	if (defined $print) {
	    my $print_reg;
	    $shell_reg->Create("print", $print_reg);
	    my $command_reg;
	    $print_reg->Create("command", $command_reg);
	    $command_reg->SetValue("", REG_SZ, $print);
	}
    };
    warn $@ if ($@);
}

=head1 USER FUNCTIONS

=head2 get_user_name

Get current windows user.

=cut

sub get_user_name {
    _get_api_function("GetUserName");
    return unless $API_FUNC{GetUserName};
    my $max = 256;
    my $maxb = pack("L", $max);
    my $login = "\0"x$max;
    my $b = $API_FUNC{GetUserName}->Call($login, $maxb);
    if ($b) {
	substr($login, 0, unpack("L", $maxb)-1);
    } else {
	undef;
    }
}

=head2 is_administrator

Guess if current user has admin rights.

=cut

sub is_administrator {
    my $user_name = get_user_name();
    if (defined $user_name) {
	return $user_name =~/^(administrator|admin)$/i ? 1 : 0;
    }
    undef;
}

=head2 get_user_folder($foldertype, $public)

Get the folder path for the current user, or, if $public is set to a
true value, for the whole system. If $foldertype is not given, the
"Personal" subfolder is returned.

=cut

sub get_user_folder {
    my($foldertype, $public) = @_;
    $foldertype = 'Personal' if !defined $foldertype;
    if ($public) {
	my $common_folders =
	    { map { $_ => 1 }
	      (qw/AppData Desktop Programs Startup/, 'Start Menu')
	    };
	if (exists $common_folders->{$foldertype}) {
	    $foldertype = "Common $foldertype";
	}
    }
    my $folder;
    eval q{
        use Win32::Registry;
        my $top_hkey = ($public
			? $main::HKEY_LOCAL_MACHINE
			: $main::HKEY_CURRENT_USER);
        my($reg_key, $key_ref, $hashref);
        $reg_key = join('\\\\', qw(SOFTWARE Microsoft Windows CurrentVersion
				   Explorer), 'Shell Folders');
        return unless $top_hkey->Open($reg_key, $key_ref);
        return unless $key_ref->GetValues($hashref);
        $folder = $hashref->{$foldertype}[2];
    };
    warn $@ if $@;
    $folder;
}

=head1 WWW AND NET FUNCTIONS

=head2 lwp_auto_proxy($lwp_user_agent)

Set the proxy for a LWP::UserAgent object (similar to the unix-centric
env_proxy method). Uses the Internet Explorer proxy setting.

=cut

sub lwp_auto_proxy {
    my $lwp_user_agent = shift;

    my $proxy_server;
    my $proxy_enable = 0;
    my $proxy_override;

    eval q{
        use Win32::Registry;
	my($reg_key, $key_ref, $hashref);
        $reg_key = join('\\\\', qw/Software Microsoft Windows CurrentVersion/,
			'Internet Settings');
	if ($main::HKEY_CURRENT_USER->Open($reg_key, $key_ref) &&
	    $key_ref->GetValues($hashref)) {
	    $proxy_enable   = $hashref->{"ProxyEnable"}[2];
	    $proxy_server   = $hashref->{"ProxyServer"}[2];
	    $proxy_override = $hashref->{"ProxyOverride"}[2];
	}
    };

    warn "Proxy settings from registry:
  enable=$proxy_enable server=$proxy_server override=$proxy_override\n"
	if $DEBUG;

    if ($proxy_enable) {
	# It seems that the following formats are possible:
	#    [http://]proxy[:port]
	# Fix this format to the one LWP uses...
	if ($proxy_server !~ m|^.*://|) {
	    $proxy_server = "http://$proxy_server/";
	}
	warn "Using <$proxy_server> as LWP proxy server setting\n"
	    if $DEBUG;
	$lwp_user_agent->proxy(['http', 'ftp'], $proxy_server);
	if (defined $proxy_override && $proxy_override eq '<local>') {
	    # XXX There is no way to say that hosts without domain portion
	    # should be no_proxied... So this is a poor excuse...
	    $lwp_user_agent->no_proxy("127.0.0.1", "localhost");
	}
    }
}

=head1 MAIL FUNCTIONS

=head2 send_mail(%args)

Send an email through MAPI or other means. Some of the following
arguments are recognized:

=over 4

=item -sender

Required. The sender who is sending the mail.

=item -passwd

The MAPI password (?)

=item -recipient

The recipient of the mail.

=item -subject

The subject of the message.

=item -body

The body text of the message.

=back

This is from Win32 FAQ. Not tested, because MAPI is not installed on
my system.

=cut

sub send_mail {
    my(%args) = @_;
    send_mapi_mail(%args);
}

sub send_mapi_mail {
    my(%args) = @_;

    # Sender's Name and Password
    #
    my $sender = $args{-sender} or die "Sender is missing";
    my $passwd = $args{-password};

    # Create a new MAPI Session
    #
    require Win32::OLE;
    my $session;
    foreach my $mapiclass ("MAPI.Session",
			   #"MSMAPI.MAPISession"
			   ) {
	$session = Win32::OLE->new($mapiclass);
	last if ($session);
    }
    if (!$session) {
        die "Could not create a new MAPI Session: " . Win32::OLE->LastError();
    }

    # Attempt to log on.
    #
    my $err = $session->Logon($sender, $passwd);
    if ($err) {
        die "Logon failed: $!, " . Win32::OLE->LastError();
    }

    # Add a new message to the Outbox.
    #
    my $msg = $session->Outbox->Messages->Add();

    # Add the recipient.
    #
    my $rcpt = $msg->Recipients->Add();
    $rcpt->{Name} = $args{-recipient} or die "Recipient is missing";
    $rcpt->Resolve();

    # Create a subject and a body.
    #
    $msg->{Subject} = $args{-subject} or die "Empty Message";
    $msg->{Text} = $args{-body};

    # Send the message and log off.
    #
    $msg->Update();
    $msg->Send(0, 0, 0);
    $session->Logoff();

    1;
}

=head1 EXPLORER FUNCTIONS

=head2 create_shortcut(%args)

Create a shortcut (a desktop link). The following arguments are recognized:

=over 4

=item -path

Path to program (required).

=item -args

Additional arguments for the program.

=item -icon

Path to the .ico icon file.

=item -name

Title of the program (required).

=item -file

Specify where to save the .lnk file. If -file is not given, the file
will be stored on the current user desktop. The filename will consist
of the -name parameter and the .lnk extension.

=item -desc

Description for the file.

=item -wd

Working directory of this file.

=item -public

If true, create a shortcut visible for all users.

=item -autostart

Create shortlink in Autostart folder.

=back

=cut

sub create_shortcut {
    my(%args) = @_;
    my $path   = delete $args{-path} || die "Missing -path parameter";
    my $args   = delete $args{-args};
    my $icon   = delete $args{-icon};
    my $name   = delete $args{-name} || die "Missing -name parameter";
    my $file   = delete $args{-file};
    my $desc   = delete $args{-desc};
    my $wd     = delete $args{-wd};
    my $public = delete $args{-public} || 0;
    my $autostart = delete $args{-autostart} || 0;

    eval q{
        use Win32::Shortcut;

	if (!defined $file) {
	    my $dir;
	    $dir = get_user_folder(($autostart ? "Startup" : "Desktop"),
				   $public);
	    if (!defined $dir) {
		die "Can't get Desktop or Startup directory";
	    }
warn $dir."\n";
	    $file = join('\\\\', $dir, "$name.lnk");
	}

        my $scut = new Win32::Shortcut;
        $scut->{Path}		   = $path;
	$scut->{Arguments}	   = $args if defined $args;
	$scut->{IconLocation}      = $icon if defined $icon;
	$scut->{Description}	   = $desc if defined $desc;
	$scut->{WorkingDirectory}  = $wd   if defined $wd;
        foreach my $key (keys %args) {
            $scut->{$key} = $args{$key};
        }
        $scut->{File} = $file;
        die "Can't save $file" if !$scut->Save;
    };
    warn $@ if ($@);
}

=head2 create_internet_shortcut(%args)

Create an internet shortcut. The following arguments are recognized:

=over 4

=item -url

URL for the shortcur (required).

=item -icon

Path to the .ico icon file.

=item -name

Title of the program (required).

Specify where to save the .lnk file. If -file is not given, the file
will be stored on the current user desktop. The filename will consist
of the -url parameter and the .lnk extension.

=item -desc

Description for the file (not used yet).

=back

=cut

sub create_internet_shortcut {
    my(%args) = @_;
    my $url   = delete $args{-url} || die "Missing -url parameter";
    my $icon  = delete $args{-icon};
    my $name  = delete $args{-name} || die "Missing -name parameter";
    my $file  = delete $args{-file};
    my $desc  = delete $args{-desc};
    my $public = delete $args{-public} || 0;

    eval q{
        if (!defined $file) {
            my $desktop = get_user_folder("Desktop", $public);
            if (!defined $desktop) {
    	        die "Can't get Desktop directory";
	    }
	    $file = join('\\\\', $desktop, "$name.url");
	}

	open(URL, ">$file") or die "Can't save $file: $!";
	print URL "[InternetShortcut]\n";
	print URL "URL=$url\n";
	if (defined $icon) {
	    print URL "IconFile=$icon\n";
	    print URL "IconIndex=0\n";
	}
	close URL;
    };
    warn $@ if ($@);
}

=head2 add_recent_doc($doc)

Add the specified document to the list of recent documents.

=cut

sub add_recent_doc {
    my $doc = shift;
    warn "try $doc";
    eval q{
        my $addtorecentdocs = _get_api_function("SHAddToRecentDocs");
	die $@ if !$addtorecentdocs;
	my $SHARD_PATH = 2;
	$doc .= "\0"; # XXX notwendig???
        $addtorecentdocs->Call($SHARD_PATH, $doc);
	warn "yeah";
    };
    warn $@ if $@;
}

=head2 create_program_group(%args)

Create a program group. Following arguments are recognized:

=over 4

=item -parent

Required. The name of the new program group.

=item -files

Required. The files to be included into the new program group. The
argument may be either a file name or an array with a number of file
names. The file names can be either a string or a hash like {-path =>
'path', -name => 'name'}. In the latter case, this hash will be used
as an argument for create_shortcut.

=item -public

If true, create a program group in the public section, not in the user
section of the start menu.

=back

=cut

sub create_program_group {
    my(%args) = @_;
    my $parent = delete $args{-parent} or die "Missing -parent parameter";
    my $files  = delete $args{-files} or die "Missing -files parameter";
    my $public = delete $args{-public} || 0;
    my @files;
    push @files, (ref $files eq 'ARRAY' ? @$files : $files);
    eval q{
	use File::Path;
	use File::Basename;
	my $progdir = get_user_folder("Programs", $public);
	die "Can't get user folder." if !$progdir;
	my $topdir  = "$progdir/$parent";
	if (!-d $topdir) {
	    mkpath([$topdir], 0, 0755);
	}
	foreach my $file (@files) {
	    my %shortcut_args;
	    if (ref $file eq 'HASH') {
		%shortcut_args = %$file;
	    } else {
		%shortcut_args = (-path => $file,
				  -name => basename($file),
				 );
	    }
	    if (exists $shortcut_args{-url}) {
	        $shortcut_args{-file} = "$topdir/$shortcut_args{-name}.url";
	        create_internet_shortcut(%shortcut_args);
	    } else {
	        $shortcut_args{-file} = "$topdir/$shortcut_args{-name}.lnk";
	        create_shortcut(%shortcut_args);
	    }
	}
    };
    warn $@ if $@;
}

=head1 FILE SYSTEM FUNCTIONS

=head2 get_cdrom_drives

Return a list of CDROM drives on the system.

=cut

sub get_cdrom_drives {
    my @drives;
    eval q{
	my $DRIVE_CDROM = 5;
	my $MAX_DOS_DRIVES = 26;
        my $getlogicaldrives = _get_api_function("GetLogicalDrives");
        my $getdrivetype     = _get_api_function("GetDriveType");
	die $@ if !$getlogicaldrives || !$getdrivetype;
        my $drives = $getlogicaldrives->Call();
	my @drive_bits = split(//, unpack("b*", pack("L", $drives))); # XXX V statt L?
	for my $i (0 .. $MAX_DOS_DRIVES-1) {
	    if ($drive_bits[$i]) {
		my $drive_name = chr($i + ord('A')) . ":";
		my $drive_type = $getdrivetype->Call($drive_name);
		push @drives, $drive_name if ($drive_type == $DRIVE_CDROM);
	    }
	}
    };
    warn $@ if $@;
    @drives;
}

=head2 path2unc($path)

Expand a normal absolute path to a UNC path.

=cut

sub path2unc {
    my $path = shift;
    if ($path =~ m|^([a-z]):[/\\](.*)|i) {
        "\\\\" . Win32::NodeName() . "\\" . $1 . "\\" . $2; 
    } else {
	$path;
    }
}

=head1 GUI FUNCTIONS

=head2 client_window_region($tk_window)

Return maximum region for a window (without borders, title bar,
taskbar area). Format is ($x, $y, $width, $height).

=cut

sub client_window_region {
    my $top = shift;

    my $SPI_GETWORKAREA = 48;
    my $SM_CYCAPTION = 4;
    #my $SM_CXBORDER = 5;
    #my $SM_CYBORDER = 6;
    #my $SM_CXEDGE = 45;
    #my $SM_CYEDGE = 46;
    my $SM_CXFRAME = 32;
    my $SM_CYFRAME = 33;


    my @extends;
    _get_api_function("SystemParametersInfo");
    _get_api_function("GetSystemMetrics");
    if (!$API_FUNC{SystemParametersInfo} ||
	!$API_FUNC{GetSystemMetrics}) {
	# guess region
	@extends = (0, 0, $top->screenwidth-24, $top->screenheight-40);
    } else {
	my $buf = "\0"x(4*4); # size of RECT structure
	my $r = $API_FUNC{SystemParametersInfo}->Call($SPI_GETWORKAREA, 0,
						      $buf, 0);
	# XXX $r überprüfen
	@extends = unpack("V4", $buf);
	$extends[2] -= ($extends[0] +
			$API_FUNC{GetSystemMetrics}->Call($SM_CXFRAME)*2);
	$extends[3] -= ($extends[1] +
			$API_FUNC{GetSystemMetrics}->Call($SM_CYFRAME)*2 +
			$API_FUNC{GetSystemMetrics}->Call($SM_CYCAPTION));
    }
    @extends;
}

=head2 screen_region($tk_window)

Return maximum screen size without taskbar area.

=cut

sub screen_region {
    my $top = shift;

    my $SPI_GETWORKAREA = 48;

    my @extends;
    _get_api_function("SystemParametersInfo");
    if (!$API_FUNC{SystemParametersInfo}) {
	# guess region
	@extends = (0, 0, $top->screenwidth, $top->screenheight-20);
    } else {
	my $buf = "\0"x(4*4); # size of RECT structure
	my $r = $API_FUNC{SystemParametersInfo}->Call($SPI_GETWORKAREA, 0,
						      $buf, 0);
	# XXX $r überprüfen
	@extends = unpack("V4", $buf);
    }
    @extends;
}

=head2 maximize($tk_window)

Maximize the window. If Win32::API is installed, then the taskbar will
not be obscured.

=cut

sub maximize {
    my $top = shift;
    my @extends = client_window_region($top);
    $top->geometry("$extends[2]x$extends[3]+$extends[0]+$extends[1]");
}

=head2 get_sys_color($what)

Return ($r,$g,$b) values from 0 to 255 for the requested system color.

=cut

sub get_sys_color {
    my $type = shift;
    my $name2number =
    {"scrollbar"	    => 0,
     "background"	    => 1,
     "activecaption"	    => 2,
     "inactivecaption"	    => 3,
     "menu"		    => 4,
     "window"		    => 5,
     "windowframe"	    => 6,
     "menutext"		    => 7,
     "windowtext"	    => 8,
     "captiontext"	    => 9,
     "activeborder"	    => 10,
     "inactiveborder"	    => 11,
     "appworkspace"	    => 12,
     "highlight"	    => 13,
     "highlighttext"	    => 14,
     "btnface"		    => 15,
     "btnshadow"	    => 16,
     "graytext"		    => 17,
     "btntext"		    => 18,
     "inactivecaptiontext"  => 19,
     "btnhighlight"	    => 20,
     "3ddkshadow"	    => 21,
     "3dlight"		    => 22,
     "infotext"		    => 23,
     "infobk"		    => 24,
    };
    my $number = $name2number->{$type};
    return unless defined $number;
    _get_api_function("GetSysColor");
    return unless $API_FUNC{GetSysColor};
    my $i = $API_FUNC{GetSysColor}->Call($number);
    my($r,$g,$b);
    $b = $i >> 16;
    $g = ($i >> 8) & 0xff;
    $r = $i & 0xff;
    ($r, $g, $b);
}

=head1 MISC FUNCTIONS

=head2 sort_cmp_hack($a,$b)

"use locale" does not work on Windows. This is a hack to be used in
sort for german umlauts.

=cut

sub sort_cmp_hack {
    my($s1, $s2) = @_;
    $s1 =~ tr/äöüß/aous/;
    $s2 =~ tr/äöüß/aous/;
    $s1 cmp $s2;
}

=head1 SEE ALSO

L<perlwin32|perlwin32>, L<Win32::API|Win32::API>,
L<Win32::OLE|Win32::OLE>, L<Win32::Registry|Win32::Registry>,
L<Win32::Process|Win32::Process>, L<Win32::DDE|Win32::DDE>,
L<Win32::Shortcut|Win32::Shortcut>, L<Tk|Tk>, L<LWP::UserAgent>.

=head1 AUTHOR

Slaven Rezic <eserte@cs.tu-berlin.de>

=head1 COPYRIGHT

Copyright (c) 1999, 2000 Slaven Rezic. All rights reserved.
This module is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

1;
