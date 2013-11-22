cpr
===

Consolidated Patch Report

A perl script that takes information from the sources like the following to create a consolidated report in the form of a Excel Spreadsheet (XLSX) with charts:

- A ServiceNow CMDB XLSX extract
- A Server Master XLS list
- Solaris PCA HTML patch report
- Windows WSUS CVS output
- Redhat CVS report

It can me modified to run with more or less sources. It processes the source information into dated files so that trending information can be charted, such as whether patching rates are improving. It also supports reporting on environments based on hostname and CMDB information.

Directory Structure
===================

The following directory structure is used to generate historical data:

	top/
	cpr.pl
	raw/download.csv         (Current Red Hat Satellite CSV)
	raw/index.html           (Current PCA HTML report)
	raw/wintel_latest.csv    (Current Windows WSUS report)
	raw/cmdb.csv             (Current CMDB extract)
	raw/master_list.xls      (Current Server Master List)
	old/download.csv.MMYYYY  (Archived Red Hat Satellite CSV)
	old/index.html.MMYYYY    (Archived PCA HTML report)
	old/wintel_latest.MMYYYY (Archived Windows WSUS report)
	cpr/sol_MM_YYYY          (Solaris Monthly Patch Information)
	cpr/sol_all              (All Solaris Monthly Patch Information)
	cpr/lin_MM_YYYY          (Linux Monthly Patch Information)
	cpr/lin_all              (All Linux Monthly Patch Information)
	cpr/win_MM_YYYY          (Windows Monthly Patch Information)
	cpr/win_all              (All Windows Monthly Patch Information)
	cpr/xpr_MM_YYYY          (All Monthly Patch Information)
	cpr/xpr_all              (All Patch Information)
	pci/hosts                A file with a list of PCI host names
	exc/hosts                Hosts to exclude
	xls/all_report.xlsx      All Report
	xls/windows_rpeort.xlsx  Windows Report
	xls/linux_report.xlsx    Linux Report
	xls/solaris_report.xlsx  Linux Report
	xls/pci_report.xlsx      PCI Report

Gathering Information
=====================

To generate a report the latest versions of information need to be copied in the relevant locations as detailed above. These files contain data information, which the script will use to process the information into date based files which it will then use to generate the report with historical/trending information.

For example:

The latest ServiceNow CMDB extract is copied to raw/cmdb.xlsx

The latest Red Hat Satellite report (eg https://rhelsat/rhn/CSVDownloadAction.do) is copied to raw/download.csv

The latest Server master list is copied to raw/master_list.xls

The latest Solaris PCA HTML report is copied to raw/index.html

The latest WSUS XLS extract is copied to raw/wintel_latest.csv

The name of these files can be modified in the script or manually imported.

PCI hosts
=========

PCI hosts are rated differently to normal hosts. Any outstanding security patches is marked as bad.

The script will try to get information from the CMDB extract regarding whether a host is a PCI host. Hosts can also be manually be added to the pci/hosts file.

For example:

	$ cat pci/hosts
	pcihost1
	pcihost2


Excluding hosts
===============

If there is a need to exclude hosts from the results, this can be done by adding the hostname to exc/hosts. Additionally a reason can be given for excluding the host, which will appear as a comment in the spreadsheet report.

	$ cat exc/hosts
	host1,Decommissioned
	host2,Upgade in progress

Generating Report
=================


The first time the report is run it will take a little longer as it imports the various files and creates historical files to process.

To generate a consolidated patch report of all Operating Systems and environments:

	$ ./cpr.pl -a
	Generating xls/all_report.xlsx

To generate a summarised consolidated patch report for everything (this will hide the worksheets for individual operating systems):

	$ ./cpr.pl -a -S
	Generating xls/summary_report.xlsx


To produce a traditional percentage based report of everything:

	$ ./cpr.pl -a -S -t
	Generating xls/summary_report.xlsx

To generate a report that includes Windows information only:

	$ ./cpr.pl -w
	Generating xls/windows_report.xlsx

To generate a report that include Linux information only:

	$ ./cpr.pl -l
	Generating xls/linux_report.xlsx

To generate a report that includes Solaris information only:
	
	$ ./cpr.pl -s
	Generating xls/solaris_report.xlsx

To generate a report that includes PCI information only:
	
	$ ./cpr.pl -p
	Generating xls/pci_report.xlsx

Required Packages for Perl Script
=================================

The following Perl modules are required:

	HTML::Tagset
	URI
	HTTP::Date
	LWP::MediaTypes
	parent
	Encode
	Test::More
	IO::HTML
	Compress::Raw::Bzip2
	Compress::Raw::Zlib
	IO::Compress::Bzip2
	Encode::Locale
	HTTP::Headers
	HTML::TokeParser
	Archive::Zip
	IO::File
	Storable
	Sub::Uplevel
	Test::Exception
	Carp::Clan
	Bit::Vector
	Date::Calc
	Pod::Escapes
	Test
	Pod::Simple
	Test::Pod
	Devel::Symdump
	Pod::Coverage
	Test::Pod::Coverage
	Test::Inter
	Excel::Writer::XLSX
	Switch
	Time::Piece
	Text::Iconv
	Spreadsheet::XLSX
	Text::CSV

