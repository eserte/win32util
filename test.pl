# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

BEGIN { $| = 1; print "1..1\n"; }
END {print "not ok 1\n" unless $loaded;}
BEGIN {
    if ($^O eq 'MSWin32' || $^O eq 'cygwin') {
	eval 'use Win32Util;';
	die $@ if $@;
    } else {
	# skip all tests
	print "ok 1\n";
	$loaded = 1;
	exit;
    }
}

$loaded = 1;
print "ok 1\n";

use Getopt::Long;
my $testall;
GetOptions("testall!" => \$testall);

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

#warn join(", ", Win32Util::get_sys_color("activecaption"));
#warn Win32Util::get_user_folder("Programs", 1);
#warn Win32Util::get_user_folder("Desktop", 1);
#warn Win32Util::get_user_name();
use LWP::UserAgent;my $ua = new LWP::UserAgent;
$Win32Util::DEBUG = 1;
Win32Util::lwp_auto_proxy($ua);

if ($testall) {

    package Win32Util;

    #*start_html_viewer = \&start_html_viewer_dde;
    #*start_html_viewer = \&start_html_viewer_cmd;
    #*start_ps_viewer = \&start_ps_viewer_cmd;
    #*start_ps_print = \&start_ps_print_cmd;

    foreach my $folder_type (qw(DESKTOP
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
	warn "Folder type $folder_type: " . get_special_folder($folder_type) . "\n";
    }

    warn "Postscript viewer: " . get_ps_viewer() . "\n";
    warn "User folder: " . get_user_folder() . "\n";
    warn "Public program folder: " . get_program_folder() . "\n";
    warn "Public program start menu folder: " . get_user_folder("Programs", 1) . "\n";
    warn "All drives: " . join(", ", get_drives()) . "\n";
    warn "Net drives: " . join(", ", get_drives('remote')) . "\n";
    warn "Fixed drives: " . join(", ", get_drives('fixed')) . "\n";
    warn "Removable drives: " . join(", ", get_drives('removable')) . "\n";
    warn "CDROM drives: " . join(", ", get_drives("cdrom")) . "\n";
    warn "CDROM drives: " . join(", ", get_cdrom_drives()) . "\n";
    if (0) {
	start_ps_viewer('C:\ghost\gs4.03\tiger.ps');
	start_html_viewer('c:\users\slaven\bbbike-devel\bbbike.html');
	start_mail_composer('mailto:slaven.rezic@berlin.de');
	send_mail(-sender => 'eserte@cs.tu-berlin.de',
		  -recipient => 'eserte@192.168.1.1',
		  -subject => 'Eine Test-Mail mit MAPI',
		  -body => "jfirejreg  ger\ngfhuefheirgre\nTest 1.2.3.4.....\n\ngruss slaven\n");
    }
    warn "Disable DOS box close button\n";
    Win32Util::disable_dosbox_close_button();
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


