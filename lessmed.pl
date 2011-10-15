#!/usr/bin/perl

use strict;
use Cwd;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);

$Win32::OLE::Warn = 3;

my ($sec,$min,$hour,$mday,$mon,$year,$wday, $yday,$isdst)=localtime(time);
my $timestamp = sprintf("%4d-%02d-%02d %02d:%02d:%02d", $year+1900,$mon+1,$mday,$hour,$min,$sec);
my $date = sprintf("%d\/%d\/%4d", $mon+1,$mday,$year+1900);

my $medicine = $ARGV[0];
$medicine =~ s/%([0-9A-Fa-f]{2})/chr(hex($1))/eg;

# log arguments
open(TXT, ">>", getcwd . "/apps/lessmed/lessmed.txt") || die $!;  
  print TXT join(", ", $timestamp, $ARGV[0], $medicine, "\n");
close(TXT);

if ($medicine !~ /^$/) {
  my $match = 0;
  my @matches;
  # open Excel file, look for matching medicines
  my $file = getcwd . "/apps/lessmed/lessmed.xlsx";
  my $xl = Win32::OLE->GetActiveObject("Excel.Application") || Win32::OLE->new('Excel.Application', 'Quit');
  $xl->{DisplayAlerts} = 0;
  my $wb = $xl->Workbooks->Open($file);

  my $ws = $wb->Worksheets("lessmed");
  my $rows = $ws->UsedRange->Find({What=>"*", SearchDirection=>xlPrevious, SearchOrder=>xlByRows})->{Row};
  foreach my $row (1..$rows) {
    if ($medicine eq $row || $ws->Range("A".$row)->{Value} =~ /^$medicine/) {
      push @matches, [($row, $ws->Range("A".$row)->{Value}, $ws->Range("B".$row)->{Value}, $ws->Range("C".$row)->{Value}, $ws->Range("D".$row)->{Value})];
      $match++;
    }
  }
  if ($match == 1) {
    print $matches[0][1] . ": " . join (". ", $matches[0][2], $matches[0][3], $matches[0][4]);
  } elsif ($match > 1) {
    print $match . " matches found for '" . $medicine . "'. Reply with ";
    for (my $i = 0; $i < $match; $i++) {
      print "MED " . $matches[$i][0] . " for " . $matches[$i][1];
      if ($i < $match - 2) {
        print ", ";
      } elsif ($i == $match - 2) {
        print " or ";
      }
    }
    print ".";
  } else {
    print "We're sorry, but no matches were found for '" . $medicine . "'.";
  }
} else {
  print "You must supply the full or partial name of a medicine to retrieve a result."
}

exit(0);