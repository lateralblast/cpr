#!/usr/bin/env perl

# Name:         cpr.pl
# Version:      0.6.1
# Release:      1
# License:      CC BY-NC-SA 4.0 (Creative Commons Attribution-NonCommercial-ShareAlike)
#               https://creativecommons.org/licenses/by-nc-sa/4.0/legalcode
# Group:        Reporting
# Source:       N/A
# URL:          https://github.com/richardatlateralblast/cpr
# Distribution: Solaris/Linux
# Vendor:       Lateral Blast
# Packager:     Richard Spindler <richard@lateralblast.com.au>
# Description:  Script to produce a consolidated patch report

use strict;
use warnings;

# Install any required modules that are missing
# This runs at compile time, before the modules below are loaded
# It uses cpanm if available, otherwise cpan, and needs network access
# (and a compiler for XS modules such as Text::Iconv)

BEGIN {
  # Module to check => CPAN distribution/module to install
  my %required=(
    "Spreadsheet::ParseExcel" => "Spreadsheet::ParseExcel",
    "HTML::TokeParser"        => "HTML::Parser",
    "Excel::Writer::XLSX"     => "Excel::Writer::XLSX",
    "Text::Iconv"             => "Text::Iconv",
    "Spreadsheet::XLSX"       => "Spreadsheet::XLSX",
    "Text::CSV"               => "Text::CSV",
  );
  my @missing;
  my $module;
  my $installer;
  my @command;
  # Pick up modules from a previous non-root install
  my $local_lib="$ENV{HOME}/perl5";
  foreach my $dir ("$local_lib/lib/perl5") {
    if ((-d $dir)&&(!grep { $_ eq $dir } @INC)) {
      unshift(@INC,$dir);
    }
  }
  foreach $module (sort(keys(%required))) {
    (my $file="$module.pm")=~s{::}{/}g;
    if (!eval { require $file; 1 }) {
      push(@missing,$module);
    }
  }
  if (@missing) {
    print "Missing Perl modules: ".join(", ",@missing)."\n";
    ($installer)=grep { -x "$_/cpanm" } split(/:/,$ENV{PATH});
    if ($installer) {
      @command=("$installer/cpanm","--notest");
    }
    else {
      ($installer)=grep { -x "$_/cpan" } split(/:/,$ENV{PATH});
      if (!$installer) {
        die "Neither cpanm nor cpan found, please install: ".join(" ",map { $required{$_} } @missing)."\n";
      }
      @command=("$installer/cpan","-T");
      $ENV{PERL_MM_USE_DEFAULT}=1;
    }
    foreach $module (@missing) {
      print "Installing $required{$module}\n";
      system(@command,$required{$module});
    }
    # A non-root install may have gone into ~/perl5
    foreach my $dir ("$local_lib/lib/perl5") {
      if ((-d $dir)&&(!grep { $_ eq $dir } @INC)) {
        unshift(@INC,$dir);
      }
    }
    my @failed;
    foreach $module (@missing) {
      (my $file="$module.pm")=~s{::}{/}g;
      if (!eval { require $file; 1 }) {
        push(@failed,$required{$module});
      }
    }
    if (@failed) {
      die "Unable to install: ".join(", ",@failed)."\nPlease install them manually, e.g. cpanm --installdeps .\n";
    }
  }
}

use Spreadsheet::ParseExcel;
use Getopt::Std;
use File::Basename;
use File::Copy;
use POSIX qw(strftime);
use HTML::TokeParser;
use Excel::Writer::XLSX;
use Time::Piece;
use Text::Iconv;
use Spreadsheet::XLSX;
use Text::CSV;

# Script variables

my $script_name=$0;
my $script_version=get_script_version();
my $options="achlpstwvSVi:L:P:";
my %option;
my $logo_img="img/company_name.png";
my $company="Company Name Pty Ltd";
my $title="Consolidated Patch Report";
my $author="Perl Script";
my $pca_url="";

# Declare spreadsheet variables

my $workbook;
my $worksheet;
my $xlsx_file;

# Declare directory variables

my $raw_dir="raw";
my $old_dir="old";
my $cpr_dir="cpr";
my $pci_dir="pci";
my $xls_dir="xls";
my $exc_dir="exc";

# PCI hosts file and array

my $pci_hosts_file="$pci_dir/hosts";
my @pci_hosts;

# Exclude file and array

my $exc_hosts_file="$exc_dir/hosts";
my @exc_hosts;

# CMD file and array

my $cmdb_file="$raw_dir/cmdb.xlsx";
my @cmdb_list;

# Master server list and array

my $master_file="$raw_dir/master_list.xls";
my @master_list;

# Arrays for data from files

my @file_data;
my @all_rhel_data;
my @all_pca_data;
my @all_wsus_data;
my @all_cpr_data;
my $run_date;
my @patch_info;
my %patch_seen;

# OS and Environments

my @os_names=("Windows", "Linux", "Solaris");
my @env_names=("PCI", "Prod", "Dev", "Test", "Unknown");

# Rules for working out an environment, first matching rule wins

my @master_env_rules=(
  [ qr/PROD/, "Prod" ],
  [ qr/TEST/, "Test" ],
  [ qr/DEV/,  "Dev" ],
);

my @host_env_rules=(
  [ qr/prod|p[0-9]|o[0-9]|cd[0-9]|wp/,   "Prod" ],
  [ qr/ora[0-9]|apps[0-9]|appdb[0-9]/,   "Prod" ],
  [ qr/fin[0-9]|mon[0-9]|ninja|dp[0-9]/, "Prod" ],
  [ qr/cvs[0-9]|tax[0-9]/,               "Prod" ],
  [ qr/dev|da[0-9]|do[0-9]|dd[0-9]/,     "Dev" ],
  [ qr/lab[0-9]|pd[0-9]/,                "Dev" ],
  [ qr/test|ta[0-9]|to[0-9]|wu/,         "Test" ],
);

my @cmdb_env_rules=(
  [ qr/\bdr\b|(?<!non-)(?<!non )prod/, "Prod" ],
  [ qr/dev/,                           "Dev" ],
  [ qr/test/,                          "Test" ],
);

# Rules for working out a worksheet name prefix from a file name

my @worksheet_prefix_rules=(
  [ qr/xpr_/,      "All Platforms" ],
  [ qr/win_all/,   "All Windows" ],
  [ qr/lin_all/,   "All Linux" ],
  [ qr/sol_all/,   "All Solaris" ],
  [ qr/pci_all/,   "All PCI" ],
  [ qr/win_[0-9]/, "Windows" ],
  [ qr/lin_[0-9]/, "Linux" ],
  [ qr/sol_[0-9]/, "Solaris" ],
  [ qr/pci_[0-9]/, "PCI" ],
);

# Headers

my @headers=("Hostname", "O.S.", "Env.", "Miss", "Date");

# Set some watermarks

my $low_wm=5;
my $pci_wm=0;
my $percent_wm=90;

# Solaris PCA defaults

my %pca_data;
my @pca_data;
my $pca_html="raw/index.html";

# Redhat Satellite defaults

my @rhel_data;
my $rhel_csv="raw/download.csv";

# Windows WSUS defaults

my @wsus_data;
my $wsus_csv="raw/wintel_latest.csv";

if ($#ARGV == -1) {
  print_usage();
  exit;
}
else {
  getopts($options,\%option);
}

# If given -h print usage

if ($option{'h'}) {
  print_usage();
  exit;
}

# Print script version

if ($option{'V'}) {
  print_version();
  exit;
}

#
# Reset watermarks if given command line options
#

if ($option{'L'}) {
  $low_wm=$option{'L'};
}

if ($option{'P'}) {
  $percent_wm=$option{'P'};
}

#
# If we are not asking for help or version
# check the local environment
#

check_local_env();

sub get_script_version {
  my $version="";
  my $file_handle;
  my $line;
  if (open($file_handle,"<",$0)) {
    while ($line=<$file_handle>) {
      if ($line=~/^#\s*Version:\s*(\S+)/) {
        $version=$1;
        last;
      }
    }
    close($file_handle);
  }
  return("$version\n");
}

sub print_version {
  print "$script_version";
  return;
}

#
# Import a file into history
#

if ($option{'i'}) {
  if (! -e "$option{'i'}") {
    print "File: $option{'i'} does not exist\n";
    exit;
  }
  if ($option{'v'}) {
    print "Importing $option{'i'}\n";
  }
  if ($option{'s'}) {
    $pca_html=$option{'i'};
    historical_pca_data();
  }
  if ($option{'l'}) {
    $rhel_csv=$option{'i'};
    historical_rhel_data();
  }
  if ($option{'w'}) {
    $wsus_csv=$option{'i'};
    historical_wsus_data();
  }
  merge_cpr_data();
  exit;
}

#
# Clear out consolidated reports
#

if (($option{'a'})||($option{'p'})) {
  clean_up_files();
}

#
# Process Red Hat CSV
#

if (($option{'l'})||($option{'a'})||($option{'p'})) {
  if ($option{'c'}) {
    import_rhel_csv($rhel_csv);
    dump_rhel_data();
  }
  else {
    historical_rhel_data();
    merge_rhel_data();
  }
}

#
# Process WSUS CSV
#

if (($option{'w'})||($option{'a'})||($option{'p'})) {
  if ($option{'c'}) {
    import_wsus_data();
    dump_wsus_data();
  }
  else {
    historical_wsus_data();
    merge_wsus_data();
  }
}

#
# Process Solaris PCA HTML
#

if (($option{'s'})||($option{'a'})||($option{'p'})) {
  if ($option{'c'}) {
    import_pca_data($pca_html);
    dump_pca_data();
  }
  else {
    historical_pca_data();
    merge_pca_data();
  }
}

# If given -c only dump the data, don't generate anything

if ($option{'c'}) {
  exit;
}

#
# If given any of the following options generate a spreadsheet
#

if (($option{'a'})||($option{'l'})||($option{'s'})||($option{'w'})||($option{'p'})) {
  # Consolidated files are only regenerated (see clean_up_files) for -a and -p
  if (($option{'a'})||($option{'p'})) {
    merge_cpr_data();
  }
  generate_speadsheet();
}

#
# Clean up files
# This cleans up the consolidated files
#

sub clean_up_files {
  my $file_name;
  foreach $file_name (find_files($cpr_dir,qr/xpr/)) {
    if ($option{'v'}) {
      print "Deleting $file_name\n";
    }
    unlink($file_name) or warn "Could not delete $file_name: $!\n";
  }
  if (-e "$cpr_dir/pci_all") {
    unlink("$cpr_dir/pci_all") or warn "Could not delete $cpr_dir/pci_all: $!\n";
  }
}

#
# Get a sorted list of the files in a directory with names matching a pattern
#

sub find_files {
  my $dir_name=$_[0];
  my $pattern=$_[1];
  my $dir_handle;
  my @file_list;
  if (opendir($dir_handle,$dir_name)) {
    @file_list=sort map { "$dir_name/$_" } grep { (-f "$dir_name/$_")&&($_=~$pattern) } readdir($dir_handle);
    closedir($dir_handle);
  }
  return(@file_list);
}

#
# Check whether a host is in the PCI hosts list
# Hosts are matched exactly, not as substrings
#

sub is_pci_host {
  my $host_name=lc($_[0]);
  return(scalar(grep { $_ eq $host_name } @pci_hosts));
}

#
# Find the exclude list entry (hostname,reason) for a host
# Returns an empty string if the host is not excluded
#

sub get_exclusion {
  my $host_name=$_[0];
  my ($entry)=grep { /^\Q$host_name\E\s*,/i } @exc_hosts;
  return($entry||"");
}

#
# Check the local environment
# Create sub directories if they don't exist
#

sub check_local_env {
  my @dir_list=( $raw_dir, $old_dir, $cpr_dir, $pci_dir, $xls_dir, $exc_dir );
  my $dir_name;
  foreach $dir_name (@dir_list) {
    if (! -d "$dir_name") {
      mkdir($dir_name) or die "Could not create directory $dir_name: $!\n";
    }
  }
  if (-e "$exc_hosts_file") {
    if ($option{'v'}) {
      print "Importing Excluded hosts\n";
    }
    import_file($exc_hosts_file);
    @exc_hosts=();
    foreach my $exc_host (@file_data) {
      $exc_host=~s/^\s+|\s+$//g;
      if ($exc_host=~/[A-Za-z0-9]/) {
        push(@exc_hosts,$exc_host);
      }
    }
  }
  if (-e "$pci_hosts_file") {
    if ($option{'v'}) {
      print "Importing PCI hosts\n";
    }
    import_file($pci_hosts_file);
    @pci_hosts=();
    foreach my $pci_host (@file_data) {
      $pci_host=~s/^\s+|\s+$//g;
      if ($pci_host=~/[A-Za-z0-9]/) {
        push(@pci_hosts,lc($pci_host));
      }
    }
  }
  if (-e "$master_file") {
    if ($option{'v'}) {
      print "Importing Master Server List\n";
    }
    import_master_list();
  }
  if (-e "$cmdb_file") {
    if ($option{'v'}) {
      print "Importing CMDB List\n";
    }
    import_cmdb_list();
  }
  $run_date=strftime("%d/%m/%Y",localtime());
  return;
}

#
# This routine creates a Excel::Writer object
# Depending on the type of report, the appropriate output file name is chosen
#

sub create_speadsheet {
  my $pattern;
  if ($option{'a'}) {
    $pattern=qr/./;
    if ($option{'S'}) {
      $xlsx_file="$xls_dir/summary_report.xlsx";
    }
    else {
      $xlsx_file="$xls_dir/all_report.xlsx";
    }
  }
  if ($option{'l'}) {
    $pattern=qr/lin/;
    $xlsx_file="$xls_dir/linux_report.xlsx";
  }
  if ($option{'w'}) {
    $pattern=qr/win/;
    $xlsx_file="$xls_dir/windows_report.xlsx";
  }
  if ($option{'s'}) {
    $pattern=qr/sol/;
    $xlsx_file="$xls_dir/solaris_report.xlsx";
  }
  if ($option{'p'}) {
    $pattern=qr/pci/;
    $xlsx_file="$xls_dir/pci_report.xlsx";
  }
  $workbook=Excel::Writer::XLSX->new($xlsx_file);
  if (!$option{'v'}) {
    print "Generating $xlsx_file\n";
  }
  return($pattern);
}

#
# Print usage information if script used with -h
#

sub print_usage {
  print "\n";
  print "Usage: $script_name -[$options]\n";
  print "\n";
  print "-h: Display help/usage\n";
  print "-V: Display version\n";
  print "-l: Process Linux Red Hat Satellite patching information\n";
  print "-s: Process Solaris PCA patching information\n";
  print "-w: Process Windows WSUS patching information\n";
  print "-p: Process PCI host patching information\n";
  print "-a: Proccess all patching information\n";
  print "-c: Output current data without processing (debug)\n";
  print "-i: Input raw data from file (used with -s, -l, or -w)\n";
  print "-S: Print summarised report (Cover sheet and All Platforms)\n";
  print "-L: Set low watermark\n";
  print "-P: Set percentage watermark\n";
  print "-t: Do a traditional percentage based report\n";
  print "-v: Verbose output\n";
  print "\n";
  return;
}

#
# Import CMDB and clean data
# This uses the XLSX import module, the data is in the format below:
# Name,Class,Short description,Manufacturer,Location,OS Service Pack,
# OS Version,OS Address Width (bits),OS Domain,Operating System,Operational status
# Get the information we need and put it into an array:
# Currently we use Hostname, OS, Environment, and Description
#

sub import_cmdb_list {
  my $host_name;
  my @data;
  my $line;
  my $lc_line;
  my $junk;
  my $os_name;
  my $os_source;
  my $env_name;
  my $server_info;
  my $description;
  my $parser=Text::Iconv->new("utf-8", "windows-1251");
  my $excel=Spreadsheet::XLSX ->new($cmdb_file,$parser);
  foreach my $sheet (@{$excel -> {Worksheet}}) {
    $sheet->{MaxRow}||=$sheet->{MinRow};
    foreach my $row ($sheet->{MinRow}..$sheet->{MaxRow}) {
      $sheet->{MaxCol}||=$sheet->{MinCol};
      @data=();
      $line="";
      foreach my $col ($sheet->{MinCol}..$sheet->{MaxCol}) {
        my $cell=$sheet->{Cells}[$row][$col];
        $cell=($cell&&defined($cell->{Val}))?$cell->{Val}:"";
        $cell=~s/\n/ /g;
        $cell=~s/,/ /g;
        push(@data,$cell);
      }
      $line=join(",",@data);
      push(@file_data,$line);
    }
  }
  foreach $line (@file_data) {
    chomp($line);
    @data=split(/,/,$line);
    $description=$data[2];
    $lc_line=lc($line);
    if ($lc_line!~/decom|non-operational/) {
      $env_name="";
      $os_name="";
      if ($lc_line=~/windows|linux|solaris/) {
        # Prefer the Operating System column over the rest of the row
        $os_source=(defined($data[9])&&lc($data[9])=~/windows|linux|solaris/)?lc($data[9]):$lc_line;
        if ($lc_line=~/(?<!non-)(?<!non )prod/) {
          $env_name="Prod";
        }
        if ($lc_line=~/dev/) {
          $env_name="Dev";
        }
        if ($lc_line=~/test/) {
          $env_name="Test";
        }
        if ($os_source=~/windows/) {
          $os_name="Windows";
        }
        if ($os_source=~/linux/) {
          $os_name="Linux";
        }
        if ($os_source=~/solaris/) {
          $os_name="Solaris";
        }
        @data=split(/,/,$line);
        $host_name=$data[0];
        if ($host_name=~/\s+-/) {
          ($host_name,$junk)=split(/\s+-/,$host_name);
        }
        $host_name=~s/\s+//g;
        $host_name=~s/"//g;
        $host_name=lc($host_name);
        if ($host_name!~/migrated|template|clone|patched/) {
          if ($os_name=~/Solaris/) {
            if ($host_name=~/^au/) {
              $server_info="$host_name,$os_name,$env_name,$description";
              push(@cmdb_list,$server_info);
            }
          }
          else {
            $server_info="$host_name,$os_name,$env_name,$description";
            push(@cmdb_list,$server_info);
          }
        }
      }
    }
  }
}

#
# Parse a date string from a report
# Returns a Time::Piece object, or undef if the string is not a valid date
#

sub parse_report_date {
  my $date_string=$_[0];
  my $format=$_[1];
  $date_string=~s/^\s+|\s+$//g;
  if ($date_string eq "") {
    return;
  }
  my $date=eval { Time::Piece->strptime($date_string,$format) };
  return($date);
}

#
# Generate a file suffix
# Used for archiving files
#

sub generate_file_suffix {
  my $date_string=$_[0];
  my $month;
  my $year;
  my $prefix;
  ($prefix,$month,$year)=split("-",$date_string);
  $date_string="$month"."_"."$year";
  return($date_string);
}

#
# Get the value of a spreadsheet cell
# Empty cells are returned as undef by Spreadsheet::ParseExcel
#

sub cell_value {
  my $cell=$_[0];
  if (!$cell) {
    return("");
  }
  return($cell->value());
}

#
# Load Server Master List into an array if it exists
# Used to get more information about servers
#

sub import_master_list {
  my $host_name;
  my $row;
  my $row_min;
  my $row_max;
  my $landscape;
  my $application;
  my $server_info;
  my $sheet;
  my $input_worksheet;
  my $parser=Spreadsheet::ParseExcel->new();
  my $input_workbook=$parser->Parse($master_file);
  if (!$input_workbook) {
    warn "Could not parse $master_file: ".$parser->error()."\n";
    return;
  }
  # Worksheet number and the application column for each platform
  # Sheet 0 has the Solaris hosts, sheet 1 has the Linux hosts
  my @sheets=( [ 0, ord('U')-65 ], [ 1, ord('Y')-65 ] );
  foreach $sheet (@sheets) {
    $input_worksheet=$input_workbook->worksheet($sheet->[0]);
    if (!$input_worksheet) {
      next;
    }
    ($row_min,$row_max)=$input_worksheet->row_range();
    for $row ($row_min .. $row_max) {
      $host_name=cell_value($input_worksheet->get_cell($row,0));
      $host_name=~s/\s+//g;
      if ($host_name=~/[a-z]/) {
        $landscape=cell_value($input_worksheet->get_cell($row,4));
        $landscape=~s/\s+//g;
        $application=cell_value($input_worksheet->get_cell($row,$sheet->[1]));
        $application=~s/\n/ /g;
        $application=~s/^\s+//g;
        $application=~s/ $//g;
        $server_info="$host_name,$landscape,$application";
        push(@master_list,$server_info);
      }
    }
  }
  return;
}

#
# Determine environment
# This function is called if we are not able to determine the environment name
#

sub get_environment {
  my $host_name=$_[0];
  my $env_name="";
  my $server_info;
  my $lc_server_info;
  my $match;
  # Use the Server Master List first (last matching entry wins)
  foreach $server_info (@master_list) {
    if ($server_info=~/^\Q$host_name\E,/) {
      $match=first_match($server_info,@master_env_rules);
      if (defined($match)) {
        $env_name=$match;
      }
    }
  }
  # Otherwise try and work it out from the host name
  if ($env_name eq "") {
    $match=first_match($host_name,@host_env_rules);
    if (defined($match)) {
      $env_name=$match;
    }
  }
  # Otherwise use the CMDB information
  if ($env_name eq "") {
    foreach $server_info (@cmdb_list) {
      $lc_server_info=lc($server_info);
      if ($lc_server_info=~/^\Q$host_name\E,/) {
        $match=first_match($lc_server_info,@cmdb_env_rules);
        if (defined($match)) {
          $env_name=$match;
        }
      }
    }
  }
  if ($env_name eq "") {
    $env_name="Unknown";
  }
  if ($option{'v'}) {
    print "Setting environment for $host_name to $env_name\n";
  }
  return($env_name);
}

#
# Return the value of the first rule whose pattern matches the string
# Rules are [ pattern, value ] pairs, returns undef if none match
#

sub first_match {
  my $string=shift;
  my $rule;
  foreach $rule (@_) {
    if ($string=~$rule->[0]) {
      return($rule->[1]);
    }
  }
  return;
}

#
# Get the date of a Red Hat Satellite report
# The report has no date, so use the first check in date that appears on
# consecutive lines, as most hosts check in at the same time
#

sub get_rhel_report_date {
  my $file_name=$_[0];
  my $file_handle;
  my $line;
  my @fields;
  my $date="";
  my $last_date="";
  if (open($file_handle,"<",$file_name)) {
    while ($line=<$file_handle>) {
      chomp($line);
      @fields=split(/,/,$line,-1);
      if ((defined($fields[6]))&&($fields[6]=~/[0-9]/)) {
        ($date)=split(" ",$fields[6]);
        if ((defined($date))&&($date eq $last_date)) {
          close($file_handle);
          return($date);
        }
        $last_date=(defined($date))?$date:"";
      }
    }
    close($file_handle);
  }
  return("");
}

#
# Get the date of a Solaris PCA report
# This is fields 8 to 10 of the line containing the time zone
#

sub get_pca_report_date {
  my $file_name=$_[0];
  my $file_handle;
  my $line;
  my @fields;
  if (open($file_handle,"<",$file_name)) {
    while ($line=<$file_handle>) {
      if ($line=~/EST/) {
        @fields=split(" ",$line);
        close($file_handle);
        return(join(" ",grep { defined($_) } @fields[7..9]));
      }
    }
    close($file_handle);
  }
  return("");
}

#
# Get the date of a Windows WSUS report
# This is the first field of the last line that contains a number
#

sub get_wsus_report_date {
  my $file_name=$_[0];
  my $file_handle;
  my $line;
  my $last_line="";
  if (open($file_handle,"<",$file_name)) {
    while ($line=<$file_handle>) {
      if ($line=~/[0-9]/) {
        $last_line=$line;
      }
    }
    close($file_handle);
  }
  my ($date)=split(/,/,$last_line);
  return(defined($date)?$date:"");
}

#
# Generate historical Red Hat Satellite file
# Extracts date from report if a historical file doesn't exist
# it creates one e.g. patform/platform_mm_yyyy
#

sub historical_rhel_data {
  my $date_string;
  my $date_file;
  my $platform="Linux";
  my $line;
  my @data;
  my $date;
  my $host_name;
  my $critical;
  my $env_name;
  my $pci_file;
  my $cmdb_host;
  my $cmdb_line;
  my $cmdb_os;
  my $cmdb_env;
  my $rhel_test;
  my $description;
  my $output;
  my $pci_output;
  if (-e "$rhel_csv") {
    # Get date by looking for common check in time
    $date_string=get_rhel_report_date($rhel_csv);
    $date_string=parse_report_date($date_string,"%m/%d/%y");
    if (!$date_string) {
      print "Unable to determine report date from $rhel_csv\n";
      return;
    }
    $date_string=$date_string->dmy;
    $date_string=generate_file_suffix($date_string);
    $date_file="$old_dir/".basename($rhel_csv)."_"."$date_string";
    # Make a dated copy of the raw data
    if (!-e "$date_file") {
      if ($option{'v'}) {
        print "Archiving $rhel_csv to $date_file\n";
      }
      copy($rhel_csv,$date_file) or warn "Could not archive $rhel_csv to $date_file: $!\n";
    }
    # Process raw date into a standard dated file
    $date_file="$cpr_dir/lin_$date_string";
    $pci_file="$cpr_dir/pci_$date_string";
    if (! -e "$date_file") {
      open($output,">",$date_file) or die "Could not open $date_file: $!\n";
      open($pci_output,">>",$pci_file) or die "Could not open $pci_file: $!\n";
      import_rhel_csv();
      foreach $cmdb_line (@cmdb_list) {
        if ($cmdb_line=~/Linux/) {
          $rhel_test=0;
          ($cmdb_host,$cmdb_os,$cmdb_env,$description)=split(/,/,$cmdb_line);
          foreach $line (@rhel_data) {
            $env_name="";
            #  "Hostname", "Platform", "Environment", Missing", "Date"
            chomp($line);
            if ($line!~/Security Errata/) {
              @data=split(",",$line);
              $host_name=$data[0];
              # some hosts appear with quotations, strip them
              $host_name=~s/"//g;
              # some hosts use FQHN, strip off domain name
              if ($host_name=~/\./) {
                ($host_name)=split /\./, $host_name;
              }
              $host_name=lc($host_name);
              if ($host_name eq $cmdb_host) {
                $rhel_test=1;
                # get the number of critical patches outstanding
                $critical=$data[2];
                $date=$data[6];
                # Adjust US date to AU date
                @data=split(" ",$date);
                $date=$data[0];
                $date=Time::Piece->strptime($date,"%m/%d/%y");
                $date=$date->dmy("/");
                if (is_pci_host($host_name)) {
                  $env_name="PCI";
                  print {$pci_output} "$host_name,$platform,$env_name,$critical,,$date\n";
                }
                if ($env_name!~/[A-Za-z]/) {
                  $env_name=get_environment($host_name);
                }
                print {$output} "$host_name,$platform,$env_name,$critical,,$date\n";
              }
            }
          }
          if ($rhel_test == 0) {
            $env_name="";
            $critical="N/A";
            $date="None";
            if (is_pci_host($cmdb_host)) {
              $env_name="PCI";
              print {$pci_output} "$cmdb_host,$platform,$env_name,$critical,,$run_date\n";
            }
            if ($env_name!~/[A-Za-z]/) {
              $env_name=get_environment($cmdb_host);
            }
            print {$output} "$cmdb_host,$platform,$env_name,$critical,,$run_date\n";
          }
        }
      }
      close($output);
      close($pci_output);
    }
  }
  return;
}

#
# Generate historical Solaris PCA file
# Extracts date from report if a historical file doesn't exist
# it creates one e.g. patform/platform_mm_yyyy
#

sub historical_pca_data {
  my $date_string;
  my $date_file;
  my $platform="Solaris";
  my $line;
  my @data;
  my $date;
  my $host_name;
  my $critical;
  my $env_name;
  my $pci_file;
  my $cmdb_host;
  my $cmdb_line;
  my $cmdb_os;
  my $cmdb_env;
  my $pca_test;
  my $description;
  my $host_id;
  my $output;
  my $pci_output;
  if (-e "$pca_html") {
    $date_string=get_pca_report_date($pca_html);
    $date_string=parse_report_date($date_string,"%d %B %Y");
    if (!$date_string) {
      print "Unable to determine report date from $pca_html\n";
      return;
    }
    $date_string=$date_string->dmy;
    $date_string=generate_file_suffix($date_string);
    $date_file="$old_dir/".basename($pca_html)."_"."$date_string";
    # Make a dated copy of the raw data
    if (!-e "$date_file") {
      if ($option{'v'}) {
        print "Archiving $pca_html to $date_file\n";
      }
      copy($pca_html,$date_file) or warn "Could not archive $pca_html to $date_file: $!\n";
    }
    # Process raw date into a standard dated file
    $date_file="$cpr_dir/sol_$date_string";
    $pci_file="$cpr_dir/pci_$date_string";
    if (!-e "$date_file") {
      open($output,">",$date_file) or die "Could not open $date_file: $!\n";
      open($pci_output,">>",$pci_file) or die "Could not open $pci_file: $!\n";
      import_pca_data();
      foreach $cmdb_line (@cmdb_list) {
        if ($cmdb_line=~/Solaris/) {
          $pca_test=0;
          ($cmdb_host,$cmdb_os,$cmdb_env,$description)=split(/,/,$cmdb_line);
          foreach $line (@pca_data) {
            $env_name="";
            #  "Hostname", "Platform", "Environment", Missing", "Date"
            chomp($line);
            # get the number of critical patches outstanding
            @data=split(",",$line);
            $critical=$data[-3];
            $date=$data[2];
            $host_id=$data[3];
            $host_name=$data[0];
            # some hosts appear with quotations, strip them
            $host_name=~s/"//g;
            # some hosts use FQHN, strip off domain name
            if ($host_name=~/\./) {
              ($host_name)=split /\./, $host_name;
            }
            $host_name=lc($host_name);
            if ($host_name eq $cmdb_host) {
              $pca_test=1;
              # Fix date
              @data=split(/\./,$date);
              $date="$data[2]/$data[1]/$data[0]";
              if (is_pci_host($host_name)) {
                $env_name="PCI";
                print {$pci_output} "$host_name,$platform,$env_name,$critical,$host_id,$date\n";
              }
              if ($env_name!~/[A-Za-z]/) {
                $env_name=get_environment($host_name);
              }
              print {$output} "$host_name,$platform,$env_name,$critical,$host_id,$date\n";
            }
          }
          if ($pca_test == 0) {
            $env_name="";
            $critical="N/A";
            $date="None";
            $host_id="";
            if (is_pci_host($cmdb_host)) {
              $env_name="PCI";
              print {$pci_output} "$cmdb_host,$platform,$env_name,$critical,$host_id,$run_date\n";
            }
            if ($env_name!~/[A-Za-z]/) {
              $env_name=get_environment($cmdb_host);
            }
            print {$output} "$cmdb_host,$platform,$env_name,$critical,$host_id,$run_date\n";
          }
        }
      }
      close($output);
      close($pci_output);
    }
  }
  return;
}

#
# Generate historical Windows WSUS file
# Extracts date from report if a historical file doesn't exist
# it creates one e.g. patform/platform_mm_yyyy
#

sub historical_wsus_data {
  my $date_string;
  my $date_file;
  my $platform="Windows";
  my $line;
  my @data;
  my $date;
  my $host_name;
  my $critical;
  my $env_name;
  my $pci_file;
  my $cmdb_host;
  my $cmdb_line;
  my $cmdb_os;
  my $cmdb_env;
  my $wsus_test;
  my $description;
  my $patch_info;
  my $csv=Text::CSV->new({quote_char => '"'});
  my $output;
  my $pci_output;
  if (-e "$wsus_csv") {
    $date_string=get_wsus_report_date($wsus_csv);
    $date_string=parse_report_date($date_string,"%d/%m/%Y");
    if (!$date_string) {
      print "Unable to determine report date from $wsus_csv\n";
      return;
    }
    $date_string=$date_string->dmy;
    $date_string=generate_file_suffix($date_string);
    $date_file="$old_dir/".basename($wsus_csv)."_"."$date_string";
    # Make a dated copy of the raw data
    if (!-e "$date_file") {
      if ($option{'v'}) {
        print "Archiving $wsus_csv to $date_file\n";
      }
      copy($wsus_csv,$date_file) or warn "Could not archive $wsus_csv to $date_file: $!\n";
    }
    # Process raw date into a standard dated file
    $date_file="$cpr_dir/win_$date_string";
    $pci_file="$cpr_dir/pci_$date_string";
    if (!-e "$date_file") {
      open($output,">",$date_file) or die "Could not open $date_file: $!\n";
      open($pci_output,">>",$pci_file) or die "Could not open $pci_file: $!\n";
      import_wsus_data();
      foreach $cmdb_line (@cmdb_list) {
        if ($cmdb_line=~/Windows/) {
          $wsus_test=0;
          ($cmdb_host,$cmdb_os,$cmdb_env,$description)=split(/,/,$cmdb_line);
          foreach $line (@wsus_data) {
            $env_name="";
            if ($line!~/Client/) {
              #  "Hostname", "Platform", "Missing", "Date", "PCI"
              # get the number of critical patches outstanding
              $csv->parse($line);
              @data=$csv->fields();
              $date=$data[0];
              $host_name=$data[1];
              $critical=$data[2];
              $patch_info=$data[5];
              $patch_info=~s/,/ /g;
              # some hosts appear with quotations, strip them
              $host_name=~s/"//g;
              # some hosts use FQHN, strip off domain name
              if ($host_name=~/\./) {
                ($host_name)=split /\./, $host_name;
              }
              $host_name=lc($host_name);
              if ($host_name eq $cmdb_host) {
                $wsus_test=1;
                if (is_pci_host($host_name)) {
                  $env_name="PCI";
                  print {$pci_output} "$host_name,$platform,$env_name,$critical,$patch_info,$date\n";
                }
                if ($env_name!~/[A-Za-z]/) {
                  $env_name=get_environment($host_name);
                }
                print {$output} "$host_name,$platform,$env_name,$critical,$patch_info,$date\n";
              }
            }
          }
          if ($wsus_test == 0) {
            $env_name="";
            $critical="N/A";
            $date="None";
            $patch_info="";
            if (is_pci_host($cmdb_host)) {
              $env_name="PCI";
              print {$pci_output} "$cmdb_host,$platform,$env_name,$critical,$patch_info,$run_date\n";
            }
            if ($env_name!~/[A-Za-z]/) {
              $env_name=get_environment($cmdb_host);
            }
            print {$output} "$cmdb_host,$platform,$env_name,$critical,$patch_info,$run_date\n";
          }
        }
      }
      close($output);
      close($pci_output);
    }
  }
  return;
}

#
# Import data from file
# Generic routine to read a file into an array
#

sub import_file {
  my $file_name=$_[0];
  my $file_handle;
  @file_data=();
  if (-e "$file_name") {
    open($file_handle,"<",$file_name) or die "Could not open $file_name: $!\n";
    @file_data=<$file_handle>;
    close($file_handle);
  }
  return;
}

#
# Write lines to a file
# The mode is ">" to overwrite the file or ">>" to append to it
#

sub write_lines {
  my $file_name=shift;
  my $mode=shift;
  my $file_handle;
  open($file_handle,$mode,$file_name) or die "Could not open $file_name: $!\n";
  print {$file_handle} @_;
  close($file_handle) or die "Could not write $file_name: $!\n";
  return;
}

#
# Import Red Hat Satellite data
#

sub import_rhel_csv {
  import_file($rhel_csv);
  @rhel_data=@file_data;
  return;
}

#
# Dump Red Hat data to STDOUT
#

sub dump_rhel_data {
  my $line;
  foreach $line (@rhel_data) {
    print "$line";
  }
  return;
}

#
# Import data from Redhat Satellite CSV file
#

sub import_wsus_data {
  import_file($wsus_csv);
  @wsus_data=@file_data;
  return;
}

#
# Dump Red Hat data to STDOUT
#

sub dump_wsus_data {
  my $line;
  foreach $line (@wsus_data) {
    print "$line";
  }
  return;
}

#
# Merge all Linux data into one array and output to file
# This routine creates monthly and consolidated files in the cpr directory
#

sub merge_rhel_data {
  my $file_name;
  my @file_list;
  my $output_file;
  # Get list of Linux historical files and build into a single array
  @all_rhel_data=();
  @file_list=find_files($cpr_dir,qr/^lin.*[0-9]/);
  foreach $file_name (@file_list) {
    import_file($file_name);
    @all_rhel_data=(@all_rhel_data,@file_data);
    if (($option{'a'})||($option{'p'})) {
      $output_file=$file_name;
      $output_file=~s/lin/xpr/g;
      write_lines($output_file,">>",@file_data);
    }
  }
  # Create a file with all the Linux historical data
  $output_file="$cpr_dir/lin_all";
  write_lines($output_file,">",@all_rhel_data);
  return;
}

#
# Merge all Windows WSUS data into one array and output to file
# This routine creates monthly and consolidated files in the cpr directory
#

sub merge_wsus_data {
  my $file_name;
  my @file_list;
  my $output_file;
  # Get a list of Windows historical files and build into a single array
  @all_wsus_data=();
  @file_list=find_files($cpr_dir,qr/^win.*[0-9]/);
  foreach $file_name (@file_list) {
    import_file($file_name);
    @all_wsus_data=(@all_wsus_data,@file_data);
    if (($option{'a'})||($option{'p'})) {
      $output_file=$file_name;
      $output_file=~s/win/xpr/g;
      write_lines($output_file,">>",@file_data);
    }
  }
  # Create a file with all the Linux historical data
  $output_file="$cpr_dir/win_all";
  write_lines($output_file,">",@all_wsus_data);
  return;
}

#
# Merge all Solaris PCA data into one array and output to file
# This routine creates monthly and consolidated files in the cpr directory
#

sub merge_pca_data {
  my $file_name;
  my @file_list;
  my $output_file;
  # Get a list of Solaris historical files and build into a single array
  @all_pca_data=();
  @file_list=find_files($cpr_dir,qr/^sol.*[0-9]/);
  foreach $file_name (@file_list) {
    import_file($file_name);
    @all_pca_data=(@all_pca_data,@file_data);
    if (($option{'a'})||($option{'p'})) {
      $output_file=$file_name;
      $output_file=~s/sol/xpr/g;
      write_lines($output_file,">>",@file_data);
    }
  }
  # Create a file with all the Linux historical data
  $output_file="$cpr_dir/sol_all";
  write_lines($output_file,">",@all_pca_data);
  return;
}

#
# Merge data for all platforms into one array and output to file
# This routine creates monthly and consolidated files in the cpr directory
#

sub merge_cpr_data {
  my $file_name;
  my @file_list;
  my $output_file;
  my $pci_file;
  my $line;
  my $host_name;
  my @data;
  # Get a list of all historical files and build into a single array
  @all_cpr_data=();
  @file_list=find_files($cpr_dir,qr/xpr.*[0-9]|[0-9].*xpr/);
  foreach $file_name (@file_list) {
    import_file($file_name);
    @all_cpr_data=(@all_cpr_data,@file_data);
  }
  # Create a file with all the Linux historical data
  $output_file="$cpr_dir/xpr_all";
  $pci_file="$cpr_dir/pci_all";
  my @pci_data;
  if (($option{'a'})||($option{'p'})) {
    foreach $line (@all_cpr_data) {
      @data=split(/,/,$line);
      $host_name=$data[0];
      if (is_pci_host($host_name)) {
        push(@pci_data,$line);
      }
    }
  }
  write_lines($output_file,">",@all_cpr_data);
  write_lines($pci_file,">",@pci_data);
}

#
# Routine to generate spreadsheet name
# This produces a worksheet name based on the consolidated
# file we are processing
#

sub generate_worksheet_name {
  my $file_name=$_[0];
  my $date_string;
  my $month;
  my $year;
  my $prefix;
  my $junk;
  my $worksheet_name;
  $prefix=first_match($file_name,@worksheet_prefix_rules);
  if (!defined($prefix)) {
    $prefix="";
  }
  if ($file_name=~/[0-9]/) {
    $date_string=$file_name;
    ($junk,$month,$year)=split(/_/,$date_string);
    $date_string="$month$year";
    $date_string=Time::Piece->strptime($date_string,"%m%Y");
    $month=$date_string->monname;
    $year=$date_string->year;
    $worksheet_name="$prefix $month $year";
  }
  else {
    $worksheet_name="$prefix";
  }
  return ($worksheet_name,$month,$year);
}

#
# Sort order for report files
# Monthly platform files come first (oldest first), then monthly consolidated
# files, then the all-time files. The consolidated worksheets are built from
# data collected while the platform worksheets are processed, so order matters.
#

sub report_file_order {
  my ($file_a,$file_b)=@_;
  my @keys;
  foreach my $file_name ($file_a,$file_b) {
    my $rank=2;
    my $month=0;
    my $year=0;
    if ($file_name=~/_(\d\d)_(\d{4})\s*$/) {
      $month=$1;
      $year=$2;
      $rank=($file_name=~/xpr_/)?1:0;
    }
    push(@keys,[$rank,$year,$month,$file_name]);
  }
  return($keys[0][0]<=>$keys[1][0] || $keys[0][1]<=>$keys[1][1] || $keys[0][2]<=>$keys[1][2] || $keys[0][3] cmp $keys[1][3]);
}

#
# Get the list of months (e.g. "Oct 2013") that have data, oldest first
#

sub get_report_periods {
  my %seen;
  my @periods;
  foreach my $line (@patch_info) {
    my $period=(split(/,/,$line))[-1];
    if ($period=~/^[A-Za-z]{3} \d{4}$/) {
      $seen{$period}=1;
    }
  }
  @periods=sort { Time::Piece->strptime($a,"%b %Y")->epoch <=> Time::Piece->strptime($b,"%b %Y")->epoch } keys(%seen);
  return(@periods);
}

#
# Create a cover sheet that includes the date the report is created
# This is also created as it is impossible to hide the first sheet,
# so having a cover sheet means we can hide any other sheet
#

sub create_cover_sheet {
  my $date_info;
  my $format;
  my $cover_os;
  $workbook->set_properties(
    title     => $title,
    author    => $author,
    company   => $company,
  );
  # Create front page
  $worksheet=$workbook->add_worksheet('Introduction');
  $worksheet->set_column(0,1,24);
  if (-e "$logo_img") {
    $worksheet->insert_image('A1',$logo_img,0,0,1.5,1.5);
  }
  $format=$workbook->add_format(border => 0, bold => 1, size => 32);
  $worksheet->set_row(0,36);
  $worksheet->write(0,1,$title,$format);
  $worksheet->set_row(2,24);
  $format=$workbook->add_format(border => 0, bold => 1, size => 20);
  if ($option{'l'}) {
    $cover_os="Linux";
  }
  if ($option{'w'}) {
    $cover_os="Windows";
  }
  if ($option{'s'}) {
    $cover_os="Solaris";
  }
  if ($option{'a'}) {
    $cover_os="Windows, Solaris and Linux";
  }
  if ($option{'p'}) {
    $cover_os="Windows, Solaris and Linux (PCI hosts)";
  }
  $worksheet->write(2,1,"Platforms: $cover_os",$format);
  $date_info=strftime("%a %b %e %H:%M:%S %Z %Y",localtime());
  $worksheet->set_row(4,24);
  $worksheet->write(4,1,$date_info,$format);
  return;
}

#
# Set worksheet defaults like column widths
#

sub set_worksheet_defaults {
  my $row=$_[0];
  my $col=$_[1];
  my $header;
  my $format;
  # Set Hostname Column width
  $worksheet->set_column(0,0,15);
  # Set Platform Column width
  $worksheet->set_column(1,1,10);
  # Set Environment Column width
  $worksheet->set_column(2,2,10);
  # Set Missing Column
  $worksheet->set_column(3,3,7);
  # Set Date Column width
  $worksheet->set_column(4,4,11);
  # Create headers for lists
  foreach $header (@headers) {
    $format=$workbook->add_format(bg_color => 'navy', color => 'white', bold => 1, border => 2, align => 'left');
    $worksheet->write($row,$col,$header,$format);
    $col++;
  }
  return($row,$col);
}

#
# Create a ket that has information regarding patch levels and colours
#

sub create_worksheet_key {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $format;
  # Create Key table
  $worksheet->set_column($key_col,$key_col,15);
  $worksheet->set_column($key_col+1,$key_col+1,10);
  $worksheet->set_column($key_col+2,$key_col+2,10);
  $worksheet->set_column($key_col+3,$key_col+3,10);
  $format=$workbook->add_format(bold => 1, border => 0, align => 'left');
  $worksheet->write($key_row,$key_col,"Key:",$format);
  $key_row++;
  $key_row++;
  $format=$workbook->add_format(bold => 1, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col,"Colour",$format);
  $worksheet->write($key_row,$key_col+1,"Patches",$format);
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"%",$format);
  $key_row++;
  # PCI information
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $format=$workbook->add_format(bg_color => 'green', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col,"Green (PCI)",$format);
  $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+1,"$pci_wm",$format);
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"100%",$format);
  $key_row++;
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $format=$workbook->add_format(bg_color => 'red', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col,"Red (PCI)",$format);
  $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+1,"> $pci_wm",$format);
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"< 100%",$format);
  $key_row++;
  # Normal environments
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $format=$workbook->add_format(bg_color => 'green', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col,"Green",$format);
  $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+1,"$low_wm or less",$format);
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"> $percent_wm%",$format);
  $key_row++;
  $format=$workbook->add_format(bg_color => 'yellow', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col,"Yellow",$format);
  $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+1,"< $low_wm",$format);
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"$percent_wm% - 100%",$format);
  $key_row++;
  $format=$workbook->add_format(bg_color => 'red', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col,"Red",$format);
  $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+1,"> $low_wm",$format);
  $format=$workbook->add_format(bold => 0, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"< $percent_wm%",$format);
  $key_row++;
  return($key_row,$key_col);
}

#
# Generate OS table based on OS and Month
#

sub generate_os_monthly_totals {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $last_patch_total=$_[2];
  my $month_counter=$_[3];
  my $date_string=$_[4];
  my $os_name=$_[5];
  my $env_name=$_[6];
  my $line;
  my $patch_host;
  my $patch_os;
  my $patch_env;
  my $patch_no;
  my $patch_check;
  my $patch_date;
  my $patch_info;
  my $os_patch_total=0;
  my $os_low_count=0;
  my $os_host_count=0;
  my $os_percent;
  my $format;
  my $code;
  my $os_info;
  my @data;
  my $trend;
  foreach $line (@patch_info) {
    @data=split(/,/,$line);
    $patch_host=$data[0];
    $patch_os=$data[1];
    $patch_env=$data[2];
    $patch_no=$data[3];
    $patch_check=$data[4];
    $patch_info=$data[5];
    $patch_date=$data[-1];
    if ($env_name=~/PCI/) {
      if (($patch_env=~/PCI/)&&($patch_date eq $date_string)) {
        if ($patch_no!~/N\/A/) {
          $os_host_count++;
          $os_patch_total=$os_patch_total+$patch_no;
          if ($patch_no == $pci_wm) {
            $os_low_count++;
          }
        }
      }
    }
    else {
      if (($patch_os=~/$os_name/)&&($patch_date eq $date_string)) {
        if ($patch_no!~/N\/A/) {
          $os_host_count++;
          $os_patch_total=$os_patch_total+$patch_no;
          if ($patch_no <= $low_wm) {
            $os_low_count++;
          }
        }
      }
    }
  }
  if ($os_host_count > 1) {
    $os_percent=$os_low_count/$os_host_count*100;
    $os_percent=sprintf("%1d",$os_percent);
    $format=$workbook->add_format(bg_color => 'white', border => 2);
    $os_info="$os_name ($date_string)";
    $worksheet->write($key_row,$key_col,$os_info,$format);
    $worksheet->write($key_row,$key_col+3,$os_host_count,$format);
    if ($env_name=~/PCI/) {
      if ($os_patch_total == $pci_wm) {
        $code="green";
      }
      else {
        $code="red";
      }
    }
    else {
      if ($option{'t'}) {
        if ($os_patch_total == 0) {
          $code="green";
        }
        else {
          if ($os_percent < $percent_wm) {
            $code="red";
          }
          else {
            $code="yellow";
          }
        }
      }
      else {
        if ($os_patch_total <= $low_wm*$os_host_count) {
          $code="green";
        }
        else {
          $code="yellow";
        }
      }
    }
    $format=$workbook->add_format(bg_color => $code, border => 2);
    $worksheet->write($key_row,$key_col+1,$os_patch_total,$format);
    $format=$workbook->add_format(bold => 0, bg_color => $code, border => 2, align => 'right');
    $worksheet->write($key_row,$key_col+2,"$os_percent%",$format);
    if($month_counter != 0) {
      if ($os_patch_total == $last_patch_total) {
        $trend="No change";
        $code="white";
        $month_counter++;
      }
      else {
        if ($os_patch_total < $last_patch_total) {
          $trend="Down";
          $code="green";
          $month_counter++;
        }
        else {
          if ($os_patch_total > $last_patch_total) {
            $trend="Up";
            $code="yellow";
            $month_counter++;
          }
        }
      }
    }
    else {
      $trend="N/A";
      $code="white";
      $month_counter++;
    }
    $format=$workbook->add_format(bold => 0, bg_color => $code, border => 2, align => 'center');
    $worksheet->write($key_row,$key_col+4,$trend,$format);
    $key_row++;
  }
  return($key_row,$key_col,$os_patch_total,$month_counter);
}

#
# Generate OS table based on Environment and Month
#

sub generate_env_monthly_totals {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $last_patch_total=$_[2];
  my $month_counter=$_[3];
  my $date_string=$_[4];
  my $env_name=$_[5];
  my $os_name=$_[6];
  my $line;
  my $patch_host;
  my $patch_os;
  my $patch_env;
  my $patch_no;
  my $patch_check;
  my $patch_date;
  my $env_patch_total=0;
  my $env_low_count=0;
  my $env_host_count=0;
  my $env_percent;
  my $format;
  my $code;
  my $env_info;
  my @data;
  my $trend;
  foreach $line (@patch_info) {
    @data=split(/,/,$line);
    $patch_host=$data[0];
    $patch_os=$data[1];
    $patch_env=$data[2];
    $patch_no=$data[3];
    $patch_check=$data[4];
    $patch_date=$data[-1];
    if ($os_name=~/[A-Za-z]/) {
      if ($patch_os=~/$os_name/) {
        if (($patch_env=~/$env_name/)&&($patch_date eq $date_string)) {
          if ($env_name=~/Unknown/) {
            $env_host_count++;
          }
          else {
            if ($patch_no!~/N\/A/) {
              $env_host_count++;
              $env_patch_total=$env_patch_total+$patch_no;
              if ($env_name=~/PCI/) {
                if ($patch_no == $pci_wm) {
                  $env_low_count++;
                }
              }
              else {
                if ($patch_no <= $low_wm) {
                  $env_low_count++;
                }
              }
            }
          }
        }
      }
    }
    else {
      if (($patch_env=~/$env_name/)&&($patch_date eq $date_string)) {
        if ($env_name=~/Unknown/) {
          $env_host_count++;
        }
        else {
          if ($patch_no!~/N\/A/) {
            $env_host_count++;
            $env_patch_total=$env_patch_total+$patch_no;
            if ($env_name=~/PCI/) {
              if ($patch_no == $pci_wm) {
                $env_low_count++;
              }
            }
            else {
              if ($patch_no <= $low_wm) {
                $env_low_count++;
              }
            }
          }
        }
      }
    }
  }
  if ($env_host_count > 1) {
    $env_percent=$env_low_count/$env_host_count*100;
    $env_percent=sprintf("%1d",$env_percent);
    $format=$workbook->add_format(bg_color => 'white', border => 2);
    $env_info="$env_name ($date_string)";
    $worksheet->write($key_row,$key_col,$env_info,$format);
    $worksheet->write($key_row,$key_col+3,$env_host_count,$format);
    if ($env_name=~/Unknown/) {
      $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'right');
      $worksheet->write($key_row,$key_col+1,"N/A",$format);
      $worksheet->write($key_row,$key_col+2,"N/A",$format);
      $format=$workbook->add_format(bg_color => 'white', border => 2, align => 'center');
      $worksheet->write($key_row,$key_col+4,"N/A",$format);
    }
    else {
      if ($env_name=~/PCI/) {
        if ($env_patch_total == $pci_wm) {
          $code="green";
        }
        else {
          $code="red";
        }
      }
      else {
        if ($option{'t'}) {
          if ($env_patch_total == 0) {
            $code="green";
          }
          else {
            if ($env_percent < $percent_wm) {
              $code="red";
            }
            else {
              if ($env_percent == 100) {
                $code="green"
              }
              else {
                $code="yellow";
              }
            }
          }
        }
        else {
          if ($env_patch_total <= $low_wm*$env_host_count) {
            $code="green";
          }
          else {
            $code="yellow";
          }
        }
      }
      $format=$workbook->add_format(bg_color => $code, border => 2);
      $worksheet->write($key_row,$key_col+1,$env_patch_total,$format);
      $format=$workbook->add_format(bold => 0, bg_color => $code, border => 2, align => 'right');
      $worksheet->write($key_row,$key_col+2,"$env_percent%",$format);
      if($month_counter != 0) {
        if ($env_patch_total == $last_patch_total) {
          $trend="No change";
          $code="white";
          $month_counter++;
        }
        else {
          if ($env_patch_total < $last_patch_total) {
            $trend="Down";
            $code="green";
            $month_counter++;
          }
          else {
            if ($env_patch_total > $last_patch_total) {
              $trend="Up";
              $code="yellow";
              $month_counter++;
            }
          }
        }
      }
      else {
        $trend="N/A";
        $code="white";
        $month_counter++;
      }
      $format=$workbook->add_format(bold => 0, bg_color => $code, border => 2, align => 'center');
      $worksheet->write($key_row,$key_col+4,$trend,$format);
    }
    $key_row++;
    if ($option{'v'}) {
      print "Producing Environment totals for $date_string $env_name $os_name\n";
      print "Total number of hosts for environment $env_host_count\n";
      print "Number of hosts below $low_wm is $env_low_count\n";
      print "Percentage is $env_percent\n";
    }
  }
  return($key_row,$key_col,$env_patch_total,$month_counter);
}

#
# Create header for monthly summaries
#

sub create_monthly_header {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $header_name=$_[2];
  my $format;
  $format=$workbook->add_format(bold => 1, border => 2, align => 'center');
  $worksheet->write($key_row,$key_col,"",$format);
  $worksheet->write($key_row,$key_col+1,"",$format);
  $worksheet->write($key_row,$key_col+2,"",$format);
  $worksheet->write($key_row,$key_col+3,"",$format);
  $worksheet->write($key_row,$key_col+4,"",$format);
  $key_row++;
  $format=$workbook->add_format(bold => 1, border => 2, align => 'center');
  $worksheet->write($key_row,$key_col,$header_name,$format);
  $worksheet->write($key_row,$key_col+1,"Patches",$format);
  $format=$workbook->add_format(bold => 1, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+2,"%",$format);
  $format=$workbook->add_format(bold => 1, border => 2, align => 'center');
  $worksheet->write($key_row,$key_col+3,"Hosts",$format);
  $format=$workbook->add_format(bold => 1, border => 2, align => 'center');
  $worksheet->write($key_row,$key_col+4,"Trend",$format);
  $key_row++;
  return($key_row,$key_col);
}

#
# Create header for all summary pages
#

sub create_all_header {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $header_name=$_[2];
  my $format;
  $format=$workbook->add_format(bold => 1, border => 2, align => 'center');
  $worksheet->write($key_row,$key_col,$header_name,$format);
  $worksheet->write($key_row,$key_col+1,"Month",$format);
  $worksheet->write($key_row,$key_col+2,"Patches",$format);
  $format=$workbook->add_format(bold => 1, border => 2, align => 'right');
  $worksheet->write($key_row,$key_col+3,"%",$format);
  $key_row++;
  return($key_row,$key_col);
}

#
# Create header for totals
#

sub create_totals_header {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $format;
  $key_row++;
  $format=$workbook->add_format(bold => 1, border => 0, align => 'left');
  $worksheet->write($key_row,$key_col,"Totals: (Outstanding patches and % of hosts below low water mark)",$format);
  $key_row++;
  $key_row++;
  return($key_row,$key_col);
}

#
# Create a chart and insert it into worksheet
#

sub create_chart {
  my $top_row=$_[0];
  my $end_row=$_[1];
  my $key_col=$_[2];
  my $chart_name=$_[3];
  my $grid_ref=$_[4];
  my $worksheet_name=$_[5];
  my $chart;
  if ($option{'v'}) {
    print "Inserting \"$chart_name\" chart into \"$worksheet_name\"\n";
  }
  $chart=$workbook->add_chart(type => 'column', embedded => 1);
  $chart->set_x_axis(name => $chart_name);
  $chart->set_legend(position => 'none');
  $chart->add_series(
    categories => [$worksheet_name,$top_row,$end_row,$key_col,$key_col],
    values     => [$worksheet_name,$top_row,$end_row,$key_col+1,$key_col+1],
    line       => { width => 2 },
  );
  $worksheet->insert_chart($grid_ref,$chart);
}

#
# Genereate platform totals
#

sub generate_platform_totals {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $date_string=$_[2];
  my $os_name=$_[3];
  my $worksheet_name=$_[4];
  my $section_name;
  my $env_name="";
  my $top_row;
  my $end_row;
  my $grid_ref;
  my $month_counter=0;
  my $last_patch_total;
  # Insert Platform header
  $section_name="Platforms";
  ($key_row,$key_col)=create_monthly_header($key_row,$key_col,$section_name);
  $top_row=$key_row;
  # Insert Platform information
  if ($option{'v'}) {
    print "Generating \"Platform\" totals for \"$worksheet_name\"\n";
  }
  if ($worksheet_name=~/All/) {
    if ($worksheet_name=~/Platforms/) {
      if ($worksheet_name=~/[0-9]/) {
        foreach $os_name (@os_names) {
          $month_counter=0;
          ($key_row,$key_col,$last_patch_total,$month_counter)=generate_os_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$os_name,$env_name);
        }
      }
      else {
        foreach $os_name (@os_names) {
          $month_counter=0;
          foreach $date_string (get_report_periods()) {
            ($key_row,$key_col,$last_patch_total,$month_counter)=generate_os_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$os_name,$env_name);
          }
        }
      }
    }
    else {
      if ($worksheet_name=~/PCI/) {
        $month_counter=0;
        foreach $date_string (get_report_periods()) {
          ($key_row,$key_col,$last_patch_total,$month_counter)=generate_os_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$os_name,"PCI");
        }
      }
      else {
        foreach $os_name (@os_names) {
          if ($worksheet_name=~/$os_name/) {
            $month_counter=0;
            foreach $date_string (get_report_periods()) {
              ($key_row,$key_col,$last_patch_total,$month_counter)=generate_os_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$os_name,$env_name);
            }
          }
        }
      }
    }
  }
  else {
    if ($worksheet_name=~/PCI/) {
      $month_counter=0;
      ($key_row,$key_col,$last_patch_total,$month_counter)=generate_os_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$os_name,"PCI");
    }
    else {
      $month_counter=0;
      ($key_row,$key_col,$last_patch_total,$month_counter)=generate_os_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$os_name,$env_name);
    }
  }
  # Insert OS Chart
  $end_row=$key_row-1;
  $grid_ref="M2";
  create_chart($top_row,$end_row,$key_col,$section_name,$grid_ref,$worksheet_name);
  return($key_row,$key_col);
}

#
# Generate Environment totals
#

sub generate_env_totals {
  my $key_row=$_[0];
  my $key_col=$_[1];
  my $date_string=$_[2];
  my $env_name=$_[3];
  my $worksheet_name=$_[4];
  my $section_name;
  my $os_name="";
  my $top_row;
  my $end_row;
  my $grid_ref;
  my $month_counter=0;
  my $last_patch_total;
  # Insert Environment header
  if ($option{'v'}) {
    print "Generating \"Environment\" totals for \"$worksheet_name\"\n";
  }
  $section_name="Environments";
  ($key_row,$key_col)=create_monthly_header($key_row,$key_col,$section_name);
  $top_row=$key_row;
  if ($worksheet_name=~/All/) {
    if ($worksheet_name=~/Platforms/) {
      if ($worksheet_name=~/[0-9]/) {
        foreach $env_name (@env_names) {
          $month_counter=0;
          ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$env_name,"");
        }
      }
      else {
        foreach $env_name (@env_names) {
          $month_counter=0;
          foreach $date_string (get_report_periods()) {
            ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$env_name,"");
          }
        }
      }
    }
    else {
      if ($worksheet_name=~/PCI/) {
        $month_counter=0;
        foreach $date_string (get_report_periods()) {
          ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,"PCI","");
        }
      }
      else {
        foreach $os_name (@os_names) {
          if ($worksheet_name=~/$os_name/) {
            foreach $env_name (@env_names) {
              $month_counter=0;
              foreach $date_string (get_report_periods()) {
                ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$env_name,$os_name);
              }
            }
          }
        }
      }
    }
  }
  else {
    if ($worksheet_name=~/[0-9]/) {
      if ($worksheet_name=~/PCI/) {
        $month_counter=0;
        ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,"PCI","");
      }
      else {
        foreach $os_name (@os_names) {
          if ($worksheet_name=~/$os_name/) {
            foreach $env_name (@env_names) {
              $month_counter=0;
              foreach $date_string (get_report_periods()) {
                ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$env_name,$os_name);
              }
            }
          }
        }
      }
    }
    else {
      $month_counter=0;
      ($key_row,$key_col,$last_patch_total,$month_counter)=generate_env_monthly_totals($key_row,$key_col,$last_patch_total,$month_counter,$date_string,$env_name,"");
    }
  }
  # Insert Environment charts
  $end_row=$key_row-1;
  $grid_ref="M20";
  create_chart($top_row,$end_row,$key_col,$section_name,$grid_ref,$worksheet_name);
  return($key_row,$key_col);
}

#
# Process the OS data from array created out of CSV and HTML files
# into a spreadsheet
# CSV fields:
# name,Id,Security Errata,Bug Errata,Enhancement Errata,Outdated Packages,Last Checkin,Entitlements
#

sub generate_speadsheet {
  my $line;
  my @data;
  my $row=0;
  my $col=0;
  my $item;
  my $key_col;
  my $key_row;
  my $worksheet_name;
  my $file_name;
  my $no_patches;
  my $date_string;
  my $host_name;
  my $env_name;
  my $os_name;
  my $format;
  my @file_list;
  my $file_size;
  my $check_date;
  my $counter=0;
  my $code;
  my $exclude;
  my $pattern;
  my $file_pattern;
  my $reason;
  my $comment;
  my $patch_info;
  my $report_year;
  my $period;
  $file_pattern=create_speadsheet();
  if ($option{'v'}) {
    print "Generating $xlsx_file\n";
  }
  create_cover_sheet();
  @file_list=sort { report_file_order($a,$b) } find_files($cpr_dir,$file_pattern);
  foreach $file_name (@file_list) {
    $file_size=-s $file_name;
    if (($file_name=~/[a-z]/)&&($file_size != 0)) {
      $col=0;
      $row=0;
      $key_row=1;
      $key_col=6;
      ($worksheet_name,$date_string,$report_year)=generate_worksheet_name($file_name);
      # Totals are keyed on month and year so that the same month in different years is kept apart
      $period=(defined($date_string)&&($date_string ne ""))?"$date_string $report_year":"";
      $worksheet=$workbook->add_worksheet($worksheet_name);
      if ($option{'S'}) {
        if ($worksheet_name!~/All Platforms/) {
          $worksheet->hide();
        }
      }
      $format=$workbook->add_format(bg_color => 'white', border =>0, bold => 1, align => 'center');
      @data=split(/ /,$worksheet_name);
      foreach $item (@data) {
        $worksheet->write($row,$col,$item,$format);
        $col++;
      }
      $col=0;
      $row++;
      $format=$workbook->add_format(bg_color => 'white', border => 2);
      if ($option{'v'}) {
        print "Processing $file_name to create \"$worksheet_name\"\n";
      }
      ($key_row,$key_col)=create_worksheet_key($key_row,$key_col);
      ($row,$col)=set_worksheet_defaults($row,$col);
      $format=$workbook->add_format(bg_color => 'white', fg_color => 'black', bold => 1, border => 2);
      $row++;
      import_file($file_name);
      foreach $line (@file_data) {
        $exclude=0;
        $pattern=1;
        $col=0;
        chomp($line);
        @data=split(/,/,$line);
        # Covert line into information
        $host_name=$data[0];
        $os_name=$data[1];
        $env_name=$data[2];
        $no_patches=$data[3];
        $patch_info=$data[4];
        if ($os_name=~/Solaris/) {
          if ($worksheet_name=~/[0-9]/) {
            $patch_info="$pca_url/$date_string$report_year/$host_name"."."."$patch_info"."."."html";
          }
          else {
            $patch_info="$pca_url/latest/$host_name"."."."$patch_info"."."."html";
          }
        }
        $check_date=$data[5];
        if ($env_name!~/[A-Za-z]/) {
          $env_name=get_environment($host_name);
        }
        $exclude=get_exclusion($host_name);
        if ($exclude ne "") {
          @data=split(/,/,$exclude);
          $reason=$data[1];
          $exclude=1;
          $pattern=17;
          $worksheet->write_comment($row,$col,$reason);
        }
        else {
          $exclude=0;
          $pattern=1;
          if (($comment) = grep { /^\Q$host_name\E,/ } @cmdb_list) {
            @data=split(/,/,$comment);
            $comment=$data[3];
            if ($comment=~/[A-Za-z]/) {
              $worksheet->write_comment($row,$col,$comment);
            }
          }
        }
        if ($worksheet_name!~/All/) {
          if ($exclude == 0) {
            $line="$line,$period";
            # PCI hosts appear in both the platform and PCI files, only count them once
            if (!$patch_seen{$line}++) {
              push(@patch_info,$line);
            }
          }
        }
        $format=$workbook->add_format(bg_color => 'white', border => 2, pattern => $pattern);
        # Write Hostname
        $worksheet->write($row,$col,$host_name,$format);
        $col++;
        # Write OS
        $worksheet->write($row,$col,$os_name,$format);
        $col++;
        # Write Environment
        $worksheet->write($row,$col,$env_name,$format);
        $col++;
        # Color cells according to number of patches outstaning
        if ($no_patches=~/N\/A/) {
          $code="white";
        }
        else {
          if ($env_name=~/PCI/) {
            if ($no_patches == 0) {
              $code="green"
            }
            else {
              $code="red"
            }
          }
          else {
            if ($no_patches <= $low_wm) {
              $code="green";
            }
            else {
              $code="yellow";
            }
          }
        }
        # Write Outstanding patches
        $format=$workbook->add_format(bg_color => $code, border => 2, align => 'right', pattern => $pattern);
        $worksheet->write($row,$col,$no_patches,$format);
        if ($patch_info=~/[A-Za-z]/) {
          $worksheet->write_comment($row,$col,$patch_info);
        }
        $format=$workbook->add_format(bg_color => 'white', border => 2, pattern => $pattern);
        $col++;
        $worksheet->write($row,$col,$check_date,$format);
        $row++;
      }
      # Make List filterable
      $worksheet->autofilter(1,0,$row,$col-1);
      # Create Totals Header
      ($key_row,$key_col)=create_totals_header($key_row,$key_col);
      # Generate Plaform summary
      ($key_row,$key_col)=generate_platform_totals($key_row,$key_col,$period,$os_name,$worksheet_name);
      # Generate Environment summary
      ($key_row,$key_col)=generate_env_totals($key_row,$key_col,$period,$env_name,$worksheet_name);
    }
  }
  $workbook->close();
}

#
# Process PCA HTML report
# As part of this we process the latest file and dump
# out a copy to a name file with the date so that we
# can do historical trending
#

sub import_pca_data {
  @pca_data=();
  if (! -e "$pca_html") {
    return;
  }
  my $html=HTML::TokeParser->new(shift||"$pca_html");
  my $token;
  my $text;
  my $string="";
  my $host_name="";
  my $counter=0;
  # Process every TD tag
  while ($token=$html->get_tag("td")) {
    $text=$html->get_trimmed_text("/span");
    chomp($text);
    # Hostname comes before information, so ignore first hostname so we get it's data
    if ($text=~/^au|^prd/) {
      if ($host_name=~/[a-z]/) {
        $pca_data{"$host_name"}="$string";
      }
      $host_name="$text";
      if ($counter != 0) {
        $string="";
      }
      else {
        $counter++;
      }
    }
    # Ignore headers
    if ($text!~/hostname|explorer|hostid|S\/N|Plat|Sol|KJP|Rec|Sec|Over|installed|missing|\%/) {
      if ($text!~/^au|^prd/) {
        if ($string!~/^[0-9]/) {
          $string="$string,$text";
        }
        else {
          $string="$text";
        }
      }
    }
  }
  $pca_data{"$host_name"}="$string";
  @pca_data=();
  while (($host_name,$string)=each(%pca_data)) {
    $text="$host_name,$string";
    push(@pca_data,$text);
  }
  return;
}

#
# Dump PCA data to STDOUT
#

sub dump_pca_data {
  my $line;
  foreach $line (@pca_data) {
    print "$line\n";
  }
}

