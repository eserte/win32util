# -*- perl -*-

#
# $Id: Win32Util.pm,v 1.7 1999/06/29 00:11:03 eserte Exp $
# Author: Slaven Rezic
#
# Copyright (C) 1999 Slaven Rezic. All rights reserved.
# This package is free software; you can redistribute it and/or
# modify it under the same terms as Perl itself.
#
# Mail: eserte@cs.tu-berlin.de
# WWW:  http://user.cs.tu-berlin.de/~eserte/
#

package Win32Util;

use strict;
use vars qw($DEBUG);

$DEBUG=1;

# XXX Win-Registry-Funktionen mit Hilfe von Win32::API und
# der Hilfe von der Access-Webpage nachbilden...

#*start_html_viewer = \&start_html_viewer_dde;
#*start_html_viewer = \&start_html_viewer_cmd;
#*start_ps_viewer = \&start_ps_viewer_cmd;
#*start_ps_print = \&start_ps_print_cmd;

#warn get_ps_viewer();
#start_ps_viewer('G:\ghost\gs4.03\tiger.ps');
#start_html_viewer('e:\slaven\bbbike-2.58\bbbike.html');
#start_mail_composer('eserte@onlineoffice.de');
#warn get_user_folder();
#warn get_cdrom_drives();

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

sub start_mail_composer {
    my $mailadr = shift;
    my $mailto_cmd = get_mail_composer();
    start_cmd($mailto_cmd, $mailadr);
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
sub get_mail_composer { get_reg_cmd("mailto") }
# weitere Mail-Composer-Einträge in der Registry:
# HKEY_LOCAL_MACHINE\Clients\mail\Netscape Messenger\Protocols\mailto\shell\open\command => Path to mailprg
# HKEY_LOCAL_MACHINE\Clients\mail\Pegasus Mail\Protocols\mailto\shell\open\command => Path to mailprg
# HKEY_LOCAL_MACHINE\Clients\mail\Pegasus Mail\Shell\open\command => Path to mailprg

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
        #use Win32; # XXX not needed
        my $pid;
        $r = Win32::Spawn($appname, $cmdline, $pid);
    }
    $r;
}

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

sub get_user_folder {
    my($foldertype, $public) = @_;
    $foldertype = 'Personal' if !defined $foldertype;
    my $folder;
    eval q{
        use Win32::Registry;
        my $top_hkey = ($public
			? $main::HKEY_CURRENT_MACHINE
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
    my $name = $args{-name} or die "Missing -name parameter";
    my $icon = $args{-icon};
    my $open = $args{"-open"};
    eval q{
	use Win32::Registry;
	foreach my $ext (@ext) {
	    my $ext_reg;
	    $main::HKEY_CLASSES_ROOT->Create($ext, $ext_reg);
	    $ext_reg->SetValue("", REG_SZ, $name);
	}
	my $name_reg;
	$main::HKEY_CLASSES_ROOT->Create($name, $name_reg);
	if (defined $icon) {
	    my $icon_reg;
	    $name_reg->Create("DefaultIcon", $icon_reg);
	    $icon_reg->SetValue("", REG_SZ, $icon);
	}
	if (defined $open) {
	    my $shell_reg;
	    $name_reg->Create("shell", $shell_reg);
	    my $open_reg;
	    $shell_reg->Create("open", $open_reg);
	    my $command_reg;
	    $open_reg->Create("command", $command_reg);
	    $command_reg->SetValue("", REG_SZ, $open);
	}
    };
    warn $@ if ($@);
}

# Argumente:
# -path: Pfad zum Programm (erforderlich)
# -args: zusätzliche Argumente
# -icon: Pfad zur .ico-Datei
# -name: Titel des Programms (erforderlich)
# -file: Pfad, wo die .lnk-Datei abgespeichert werden soll
#        Wenn -file nicht angegeben ist, wird der Shortcut auf dem Desktop
#        mit dem Filenamen -name .lnk abgespeichert.
sub create_shortcut {
    my(%args) = @_;
    my $path  = delete $args{-path} || die "Missing -path parameter";
    my $args  = delete $args{-args};
    my $icon  = delete $args{-icon};
    my $name  = delete $args{-name} || die "Missing -name parameter";
    my $file  = delete $args{-file};

    eval q{
        use Win32::Shortcut;

	if (!defined $file) {
	    my $desktop = get_user_folder("Desktop");
	    if (!defined $desktop) {
		die "Can't get Desktop directory";
	    }
	    $file = join('\\\\', $desktop, "$name.lnk");
	}

        my $scut = new Win32::Shortcut;
        $scut->{Path} = $path;
        $scut->{Arguments} = $args if defined $args;
        $scut->{IconLocation} = $icon if defined $icon;
        foreach my $key (keys %args) {
            $scut->{$key} = $args{$key};
        }
        $scut->{File} = $file;
        die "Can't save $file" if !$scut->Save;
    };
    warn $@ if ($@);
}

# Argumente:
# -url: URL für den Shortcut (erforderlich)
# -icon: Pfad zur .ico-Datei
# -name: Titel des Programms (erforderlich)
# -file: Pfad, wo die .url-Datei abgespeichert werden soll
#        Wenn -file nicht angegeben ist, wird der Shortcut auf dem Desktop
#        mit dem Filenamen -name .url abgespeichert.
sub create_internet_shortcut {
    my(%args) = @_;
    my $url   = delete $args{-url} || die "Missing -url parameter";
    my $icon  = delete $args{-icon};
    my $name  = delete $args{-name} || die "Missing -name parameter";
    my $file  = delete $args{-file};

    eval q{
        if (!defined $file) {
            my $desktop = get_user_folder("Desktop");
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

sub add_recent_doc {
    my $doc = shift;
    eval q{
        use Win32::API;
        my $addtorecentdocs = new Win32::API("shell32", "SHAddToRecentDocs",
					     ["I", "P"], "I");
        $addtorecentdocs->Call(2, $doc);
    };
    warn $@ if $@;
}

# Argument -files:
#   entweder nur ein Dateiname oder Array mit mehreren Dateinamen
#   jeder Dateiname kann in der Form {-path => 'path', -name => 'name'} sein,
#   wobei dieses Hash als Argument für create_shortcut verwendet wird.
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

sub get_cdrom_drives {
    my @drives;
    eval q{
	use Win32::API;
	my $DRIVE_CDROM = 5;
	my $MAX_DOS_DRIVES = 26;
        my $getlogicaldrives = new Win32::API("kernel32", "GetLogicalDrives",
					      [], "I");
        my $getdrivetype = new Win32::API("kernel32", "GetDriveType",
					  ["P"], "I");
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

# expand a normal absolute path to a UNC path
sub path2unc {
    my $path = shift;
    if ($path =~ m|^([a-z]):[/\\](.*)|i) {
        "\\\\" . Win32::NodeName() . "\\" . $1 . "\\" . $2; 
    } else {
	$path;
    }
}

1;

