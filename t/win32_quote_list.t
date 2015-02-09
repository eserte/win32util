#!/usr/bin/perl -w
# -*- cperl -*-

#
# Author: Slaven Rezic
#

use strict;
use Test::More 'no_plan';

use Win32Util ();

is(
   Win32Util::win32_quote_list('C:\Program Files\BBBike\Perl\bin\perl', '-e', 'warn q{hello, world}'),
   '"C:\Program Files\BBBike\Perl\bin\perl" -e "warn q{hello, world}"'
  );

__END__
