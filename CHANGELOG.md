# Changelog

All notable changes to this project are documented in this file.

Entries up to 0.3.0 use the original release timestamps from the former `changelog` file. Version 0.1.5 appears twice there and is kept as two entries.

## [0.6.1] - 2026-10-04

- Simplified the Server Master List import by merging the duplicated Solaris and Linux passes, dropping unused columns and handling files that cannot be parsed

## [0.6.0] - 2026-10-04

- Removed unused variables, unused globals and commented-out code, and replaced `[A-z]` character ranges with `[A-Za-z]`

## [0.5.9] - 2026-10-04

- Replaced bareword filehandles with lexical filehandles, added error checking to file opens and closes, and fixed `lin_all` and `sol_all` not being closed after writing

## [0.5.8] - 2026-10-04

- Replaced shell commands (`find`, `rm`, `mkdir`, `cp`, `date`, `cat`/`grep`/`awk`) with Perl built-ins and core modules, so filenames are no longer passed through the shell

## [0.5.7] - 2026-10-04

- Removed the deprecated `Switch` module dependency, replacing it with table-driven rules for environment and worksheet name detection

## [0.5.6] - 2026-10-04

- Enabled `use warnings` and fixed the uninitialised value warnings it exposed

## [0.5.5] - 2026-10-04

- Added automatic installation of missing required Perl modules at startup using `cpanm` or `cpan`

## [0.5.4] - 2026-10-04

- Fixed environment trends on the All Platforms sheet comparing each month against a different environment; trends are now calculated per environment

## [0.5.3] - 2026-10-04

- Fixed PCI hosts being counted twice in platform totals because they appear in both the platform and PCI data files

## [0.5.2] - 2026-10-04

- Fixed crash in `-c` mode when the Solaris PCA HTML report does not exist

## [0.5.1] - 2026-10-04

- Fixed single-platform reports (`-l`, `-w`, `-s`) regenerating consolidated files from stale data

## [0.5.0] - 2026-10-04

- Fixed missing logo image producing a warning; the logo is only inserted if the file exists

## [0.4.9] - 2026-10-04

- Fixed archived raw reports being written next to the raw files; they are now archived to `old/` as documented (README updated to the `_MM_YYYY` naming)

## [0.4.8] - 2026-10-04

- Fixed `-c` exiting after the first data source; it now dumps every requested source before exiting, and fixed the WSUS dump calling a non-existent function

## [0.4.7] - 2026-10-04

- Fixed monthly totals and trends combining the same month from different years; totals are now keyed on month and year and trended chronologically

## [0.4.6] - 2026-10-04

- Fixed multiple report options creating several workbooks, only one of which was closed

## [0.4.5] - 2026-10-04

- Fixed usage text to list the `-l` and `-p` options and drop options that are not implemented (`-r`, `-M`, `-H`)

## [0.4.4] - 2026-10-04

- Fixed percentage watermark option (`-P`) being ignored when colouring traditional reports

## [0.4.3] - 2026-10-04

- Fixed CMDB operating system detection to prefer the Operating System column, and `non-prod` being classified as Prod

## [0.4.2] - 2026-10-04

- Fixed CMDB cells containing commas shifting columns, and empty CMDB cells causing errors

## [0.4.1] - 2026-10-04

- Fixed script continuing to run after printing usage when given no arguments

## [0.4.0] - 2026-10-04

- Fixed crash when a raw report has no parsable date, and WSUS date detection when the CSV ends in a blank line

## [0.3.9] - 2026-10-04

- Fixed crash when the Server Master List contains empty cells

## [0.3.8] - 2026-10-04

- Fixed Solaris PCA report links using the current year instead of the report year

## [0.3.7] - 2026-10-04

- Fixed worksheet generation order so monthly platform sheets are processed, oldest first, before the consolidated sheets that depend on them

## [0.3.6] - 2026-10-04

- Fixed stale host ID (Solaris) and patch information (Windows) being written for hosts with no patch data

## [0.3.5] - 2026-10-04

- Fixed hostname matching for PCI hosts, excluded hosts, CMDB hosts and Satellite/PCA/WSUS hosts to match exact hostnames instead of substrings (e.g. `web1` matching `web10`)

## [0.3.4] - 2026-10-04

- Fixed CMDB Test environment being classified as Dev, and `dr` matching anywhere in a CMDB row

## [0.3.3] - 2026-10-04

- Fixed Server Master List import reading the same worksheet twice; the Linux pass now reads the second worksheet

## [0.3.2] - 2026-10-04

- Removed incorrect subtraction of 1 from platform and environment patch totals, which made PCI totals show red at zero patches

## [0.3.1] - 2026-10-04

- Fixed assignment (`=`) instead of comparison (`==`) when colouring environment percentages in traditional (`-t`) reports

## [0.3.0] - 2014-06-12

- Updated documentation and license

## [0.2.9] - 2013-11-22

- Initial GitHub release

## [0.2.8] - 2013-11-22

- Fixed number of outstanding patches in key

## [0.2.7] - 2013-10-30

- Adjusted percentages and colours

## [0.2.6] - 2013-10-23

- Fixed patch count

## [0.2.5] - 2013-10-23

- Fixed up percentage calculation and added PCI low watermark

## [0.2.4] - 2013-10-17

- Updated cover page information

## [0.2.3] - 2013-10-17

- Updated code documentations

## [0.2.2] - 2013-10-16

- Fixed individual reports

## [0.2.1] - 2013-10-15

- Added code to add patch information in comments

## [0.2.0] - 2013-10-13

- Added trending

## [0.1.9] - 2013-10-13

- Added comments from CMDB to output

## [0.1.8] - 2013-10-12

- Fixed import of HTML

## [0.1.7] - 2013-10-12

- Cleaned up code and added exclude list

## [0.1.6] - 2013-10-10

- Fixed array of months

## [0.1.5] - 2013-10-06

- Moved all spreadsheet data storage to hashes

## [0.1.5] - 2013-10-05

- Cleaned up spreadsheet creation

## [0.1.4] - 2013-10-04

- Changed watermarks based on feedback

## [0.1.3] - 2013-10-04

- Added cover page and change percentage calculation

## [0.1.2] - 2013-10-03

- Cleaned up formating

## [0.1.1] - 2013-10-02

- Added charts for monthly reports

## [0.1.0] - 2013-10-01

- Cleaned up CMDB processing and reporting

## [0.0.9] - 2013-10-01

- Fixed CMDB parsing code

## [0.0.8] - 2013-09-30

- Added code to process CMDB xlsx

## [0.0.7] - 2013-09-30

- Added code to process Master Server list

## [0.0.6] - 2013-09-30

- Added code to process CMDB extract to get environment information

## [0.0.5] - 2013-09-29

- Fixed multiple output and added PCI output

## [0.0.4] - 2013-09-27

- Fully working input and output with filters and colours

## [0.0.3] - 2013-09-25

- Added option to dump data to STDOUT

## [0.0.2] - 2013-09-25

- Working Red Hat and Solaris PCA import

## [0.0.1] - 2013-09-23

- Initial version
