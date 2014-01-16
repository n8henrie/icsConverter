# icsConverter ChangeLog

## v5.1 20140116
* Fixed a problem where numerous errors would lead to numerous popup boxes. Now exits after the first error. Or should, anyway.

## v5 20131006
* Improved logging, in large part thanks to [this](http://victorlin.me/posts/2012/08/good-logging-practice-in-python/) reader-friendly writeup.
* Added a check_dates_and_times function to help end users debug errors.
* Added some basic tests with [nose](http://nose.readthedocs.org/en/latest/).
* Made checking and allowing None values more robust.

## v4 20130924 
* Improved error handling for blank or improper headers.
* Fixed bug in reporting which header caused an error.

## v3
* Updated to read files as universal newline to avoid issues with excel-generated CSVs.
* Decided to implement a changelog because of v2.

## v2
* I actually don't remember what I changed.

## v1
* Initial release

