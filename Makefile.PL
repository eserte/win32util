use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    'NAME'		=> 'Win32Util',
    'VERSION_FROM'	=> 'Win32Util.pm', # finds $VERSION
    'PREREQ_PM'		=> {}, # e.g., Module::Name => 1.1
);

sub MY::postamble {
    my $postamble = '';
    if (defined $ENV{USER} && $ENV{USER} eq 'eserte' && $^O =~ /bsd/i && -f "../../perl.release.mk") {
	$postamble .= <<'EOF';

SF_PROJECT=	srezic
SF_PACKAGE=	Win32Util

.include "../../perl.release.mk"
.include "../../perl.cvs.mk"

EOF
    }

    $postamble;
}
