package Win32Util;

use strict;
use vars qw($DEBUG);

$DEBUG=1;

#*start_html_viewer = \&start_html_viewer_dde;
#*start_html_viewer = \&start_html_viewer_cmd;
#*start_ps_viewer = \&start_ps_viewer_cmd;
#*start_ps_print = \&start_ps_print_cmd;

#warn get_ps_viewer();
#start_ps_viewer('G:\ghost\gs4.03\tiger.ps');
#start_html_viewer('F:\perl5005\5.005\html\index.html') if $DEBUG;
#start_mail_composer('eserte@onlineoffice.de');
#warn get_user_folder();

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

sub get_html_viewer { get_reg_cmd("htmlfile") }
sub get_ps_viewer { get_reg_cmd("psfile") }
sub get_ps_print { get_reg_cmd("psfile", "print") }
sub get_mail_composer { get_reg_cmd("mailto") }

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

sub start_cmd {
    my($fullcmd, @args) = @_;
    my $r;
    eval q{
        use Win32::Process;
        use File::Basename;
        use Text::ParseWords;
        my(@words) = parse_line('\s+', 1, $fullcmd);
        my $appname = shift @words; $appname =~ s/\"//g;
        my $argstr = join(" ", @words);
        my $base = basename($appname); $base =~ s/\"//g;
        my $cmdline = $base;
        $argstr =~ s/(%(\d))/ defined($args[$2-1]) ? $args[$2-1] : "" /eg;
        $cmdline .= " $argstr";
        warn "start_cmd: " . $cmdline . "\n" if $DEBUG;
        my $proc;
        $r = Win32::Process::Create($proc, $appname, $cmdline, 0, NORMAL_PRIORITY_CLASS, ".");
    };
    $r;    
}

sub start_dde {
    my($app, $topic, $arg) = @_;
    my $r;
    eval q{
        use Win32::DDE::Client;
# XXX
$app="Netscape";# geht nur mit Netscape und nicht mit "Netscape 4.0"
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
    my($foldertype) = @_;
    $foldertype = 'Personal' if !defined $foldertype;
    my $folder;
    eval q{
        use Win32::Registry;
        my($reg_key, $key_ref, $hashref);
        $reg_key = join('\\\\', qw(SOFTWARE Microsoft Windows CurrentVersion Explorer), 'Shell Folders'); 
        return unless $main::HKEY_CURRENT_USER->Open($reg_key, $key_ref);
        return unless $key_ref->GetValues($hashref);
        $folder = $hashref->{$foldertype}[2];
    };
    warn $@ if $@;
    $folder;
}

1;

