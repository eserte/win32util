# Before `make install' is performed this script should be runnable with
# `make test'. After `make install' it should work as `perl test.pl'

######################### We start with some black magic to print on failure.

# Change 1..1 below to 1..last_test_to_print .
# (It may become useful if the test is moved to ./t subdirectory.)

BEGIN { $| = 1; print "1..1\n"; }
END {print "not ok 1\n" unless $loaded;}
BEGIN {
    if ($^O eq 'MSWin32') {
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

######################### End of black magic.

# Insert your test code below (better if it prints "ok 13"
# (correspondingly "not ok 13") depending on the success of chunk 13
# of the test code):

warn join(", ", Win32Util::get_sys_color("activecaption"));
warn Win32Util::get_user_folder("Programs", 1);
warn Win32Util::get_user_folder("Desktop", 1);
#warn Win32Util::get_user_name();

__END__

package Win32Util;

#*start_html_viewer = \&start_html_viewer_dde;
#*start_html_viewer = \&start_html_viewer_cmd;
#*start_ps_viewer = \&start_ps_viewer_cmd;
#*start_ps_print = \&start_ps_print_cmd;

warn get_ps_viewer();
start_ps_viewer('C:\ghost\gs4.03\tiger.ps');
start_html_viewer('c:\users\slaven\bbbike-devel\bbbike.html');
start_mail_composer('mailto:eserte@onlineoffice.de');
warn get_user_folder();
warn get_cdrom_drives();
send_mail(-sender => 'eserte@cs.tu-berlin.de',
	  -recipient => 'eserte@192.168.1.1',
	  -subject => 'Eine Test-Mail mit MAPI',
	  -body => "jfirejreg  ger\ngfhuefheirgre\nTest 1.2.3.4.....\n\ngruss slaven\n");
