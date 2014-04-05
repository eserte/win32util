use strict;
use Getopt::Long;
use Test::More;

BEGIN {
    if (!($^O eq 'MSWin32' || $^O eq 'cygwin')) {
	plan skip_all => 'Tests only meaningful on a Windows system';
	exit;
    }
}

plan 'no_plan';

use_ok 'Win32Util';

my $testall;
GetOptions(
	   "debug|d" => sub { $Win32Util::DEBUG = 1 },
	   "testall" => \$testall,
	  )
    or die "usage?";

{
    my @activecaption_color = Win32Util::get_sys_color("activecaption");
    ok "@activecaption_color", "Got activecaption color: @activecaption_color";
}

for my $type (qw(Programs Desktop)) {
    my $user_folder = Win32Util::get_user_folder($type, 1);
    ok $user_folder, "Got user folder for '$type': $user_folder";
}

{
    my $user_name = Win32Util::get_user_name();
    ok $user_name, "Got user name: $user_name";
}

SKIP: {
    skip "No LWP available", 1
	if !eval { require LWP::UserAgent; 1 };
    my $ua = new LWP::UserAgent;
    Win32Util::lwp_auto_proxy($ua);
    pass "Called lwp_auto_proxy";
}

foreach my $folder_type (qw(
			       DESKTOP
			       PROGRAMS
			       PERSONAL
			       FAVORITES
			       STARTUP
			       RECENT
			       SENDTO
			       STARTMENU
			       DESKTOPDIRECTORY
			       NETHOOD
			       FONTS
			       TEMPLATES
			       COMMON_STARTMENU
			       COMMON_PROGRAMS
			       COMMON_STARTUP
			       COMMON_DESKTOPDIRECTORY
			       APPDATA
			       PRINTHOOD
			       PROGRAM_FILES_COMMON
			  )) {
    my $val = Win32Util::get_special_folder($folder_type);
    ok $val, "Special folder '$folder_type': $val";
}

eval {
    Win32Util::get_special_folder('THIS_TYPE_DOES_NOT_EXIST');
};
like $@, qr{Folder type must be one of}, 'Non-existent special folder type';;

for my $def (
	     ['Postscript viewer', sub { Win32Util::get_ps_viewer() }, 1],
	     ['User folder', sub { Win32Util::get_user_folder() }],
	     ['Public program folder', sub { Win32Util::get_program_folder() }],
	     ['Public program start menu folder', sub { Win32Util::get_user_folder("Programs", 1) }],
	     ['All drives', sub { join(", ", Win32Util::get_drives()) }],
	     ['Net drives', sub { join(", ", Win32Util::get_drives('remote')) }, 1],
	     ['Fixed drives', sub { join(", ", Win32Util::get_drives('fixed')) }, 1],
	     ['Removable drives', sub { join(", ", Win32Util::get_drives('removable')) }, 1],
	     ['CDROM drives', sub { join(", ", Win32Util::get_drives("cdrom")) }, 1],
	     ['CDROM drives (via get_cdrom_drives())', sub { join(", ", Win32Util::get_cdrom_drives()) }, 1],
	    ) {
    my($label, $code, $optional) = @$def;
    my $val = $code->();
    if ($optional && !defined $val) {
	diag "$label: <nothing available>";
    } else {
	ok $val, "$label: $val";
    }
}    

if (0) {
    #*start_html_viewer = \&start_html_viewer_dde;
    #*start_html_viewer = \&start_html_viewer_cmd;
    #*start_ps_viewer = \&start_ps_viewer_cmd;
    #*start_ps_print = \&start_ps_print_cmd;
    start_ps_viewer('C:\ghost\gs4.03\tiger.ps');
    start_html_viewer('c:\users\slaven\bbbike-devel\bbbike.html');
    start_mail_composer('mailto:slaven@rezic.de');
    send_mail(-sender => 'eserte@cs.tu-berlin.de',
	      -recipient => 'eserte@192.168.1.1',
	      -subject => 'Eine Test-Mail mit MAPI',
	      -body => "jfirejreg  ger\ngfhuefheirgre\nTest 1.2.3.4.....\n\ngruss slaven\n");
}

if ($testall) {
    eval {
	Win32Util::disable_dosbox_close_button();
    };
    if ($@) {
	if ($@ =~ m{Can't locate}) {
	    diag "Prerequisites missing: $@";
	} else {
	    fail "disable_dosbox_close_button test";
	    diag $@;
	}
    } else {
	pass "disable_dosbox_close_button test";
    }
}

eval {
    require Tk;
    my $mw = MainWindow->new;
    $mw->gridRowconfigure($_, -weight => 1) for (0..2);
    $mw->gridColumnconfigure($_, -weight => 1) for (0..1);
    Tk::grid( $mw->Label(-text => q{top}), -sticky => q{n}, -columnspan => 2 );
    $mw->Label(-text => q{left})->grid(-row => 1, -column => 0, -sticky => q{w});
    $mw->Label(-text => q{right})->grid(-row => 1, -column => 1, -sticky => q{e});
    Tk::grid( $mw->Label(-text => q{bottom}, -bg => q{red}), -sticky => q{s}, -columnspan => 2 );
    $mw->update;
    Win32Util::maximize($mw);
    #$mw->update;
    $mw->after(1000, sub { $mw->destroy });
    Tk::MainLoop();
};

eval {
    require Tk;
    my $mw = MainWindow->new;
    my $l = $mw->Label(-text => "Keep on top")->pack;
    Win32Util::keep_on_top($mw, 1);
    $mw->after(2000, sub {
	$l->configure(-text => "Do not keep on top anymore");
	Win32Util::keep_on_top($mw, 0);
	$mw->after(2000, sub {
	    $mw->destroy;
	});
    });
    Tk::MainLoop();
};


