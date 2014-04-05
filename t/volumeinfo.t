#!/usr/bin/perl -w
# -*- cperl -*-

#
# Author: Slaven Rezic
#

use strict;
use Test::More;

use Win32Util;

if ($^O ne 'MSWin32') {
    plan skip_all => "Tests only meaningful on a Windows system";
    exit;
}

plan 'no_plan';

SKIP: {
    my $get_volume_information = Win32Util::_get_api_function("GetVolumeInformation");
    skip "Can't get volume info", 2
	if !$get_volume_information;
	
    ok $Win32Util::API_FUNC{GetVolumeInformation};
    is $get_volume_information, $Win32Util::API_FUNC{GetVolumeInformation};

    for my $drive (Win32Util::get_drives()) {
	my $volname = Win32Util::get_volume_name($drive."\\");
	diag "Drive=$drive VolumeName=" . (defined $volname ? $volname : '<undef>');
    }
}

__END__
