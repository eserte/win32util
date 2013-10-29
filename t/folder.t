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

plan tests => 2;

{
    my $user_folder = Win32Util::get_user_folder();
    ok $user_folder, 'get_user_folder is successful';
    ok -d $user_folder, q{... and it's a directory};
}

__END__
