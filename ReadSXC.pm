package Spreadsheet::ReadSXC;

use 5.006;
use strict;
use warnings;

require Exporter;

our @ISA = qw(Exporter);
our @EXPORT_OK = qw(read_sxc read_xml_file read_xml_string);
our $VERSION = '0.12';

use Archive::Zip;
use XML::Parser;

my %workbook;
my $table;
my $row;
my $col;
my $max_datarow;
my $max_datacol;
my $text_p_count;
my $col_count;
my @col_visible;
my %options;

sub read_sxc {
	my $zip = Archive::Zip->new(shift);
	my $content = $zip->contents('content.xml');
	my $options_ref = shift;
	my $workbook_ref = read_xml_string($content, $options_ref);
	return $workbook_ref;
}

sub read_xml_file {
	open(CONTENT, shift) or die "Cannot open file: $!";
	my $content;
	while (<CONTENT>) {
		$content .= $_;
	}
	close CONTENT;
	my $options_ref = shift;
	my $workbook_ref = read_xml_string($content, $options_ref);
	return $workbook_ref;
}

sub read_xml_string {
	my $content = shift;
	my $p = new XML::Parser(Style=>'Tree');
	my $tree = $p->parse($content);
	my $options_ref = shift;
	if ( defined $options_ref ) { %options = %{$options_ref}};
	%workbook = ();
	&collect_data($$tree[0], $$tree[1]);
	return \%workbook;
}

sub collect_data {
	my $tag = shift;
	my $node_ref = shift;
	if ( $#{$node_ref} % 2 ) { die "Bad XML parse tree: Size of array is even"; }
	my $attr_ref = shift @{$node_ref};
	my %begin_code = (
		'text:p'			=> sub	{
			if ( $text_p_count ) {
				if ( $options{'ReplaceNewlineWith'} ) {
					$workbook{$table}[$row][$col] .= $options{'ReplaceNewlineWith'};
				}
			}
			$text_p_count++;
			while ( @{$node_ref} ) {
				if ( $$node_ref[0] ) {
					&collect_data($$node_ref[0], $$node_ref[1]);
				}
				else {
					$workbook{$table}[$row][$col] .= $$node_ref[1];
				}
				splice(@{$node_ref}, 0, 2);
			}
		},
		'text:span'			=> sub	{
			while ( @{$node_ref} ) {
				if ( $$node_ref[0] ) {
					&collect_data($$node_ref[0], $$node_ref[1]);
				}
				else {
					$workbook{$table}[$row][$col] .= $$node_ref[1];
				}
				splice(@{$node_ref}, 0, 2);
			}

		},
		'office:annotation'		=> sub {
			$#{$node_ref} = -1;
		},
		'table:table-cell'		=> sub	{
			$col++;
			$workbook{$table}[$row][$col] = undef;
			$text_p_count = 0;
			if (( exists $$attr_ref{'table:date-value'} ) and ( $options{'StandardDate'} )) {
				$workbook{$table}[$row][$col] = $$attr_ref{'table:date-value'};
				$#{$node_ref} = -1;
				$text_p_count++;
			}
			else {
				while ( @{$node_ref} ) {
					&collect_data($$node_ref[0], $$node_ref[1]);
					splice(@{$node_ref}, 0, 2);
				}
			}
			if ( exists $$attr_ref{'table:number-columns-repeated'} ) {
				for (2..$$attr_ref{'table:number-columns-repeated'}) {
					$col++;
					$workbook{$table}[$row][$col] = $workbook{$table}[$row][$col - 1];
				}
			}
			if ( $text_p_count ) {
				$max_datarow = $row;
				if ( $col > $max_datacol ) { $max_datacol = $col; }
			}
		},
		'table:covered-table-cell'	=> sub	{
			$col++;
			$workbook{$table}[$row][$col] = undef;
			$text_p_count = 0;
			if ( $options{'IncludeCoveredCells'} ) {
				if (( exists $$attr_ref{'table:date-value'} ) and ( $options{'StandardDate'} )) {
					$workbook{$table}[$row][$col] = $$attr_ref{'table:date-value'};
					$#{$node_ref} = -1;
					$text_p_count++;
				}
				else {
					while ( @{$node_ref} ) {
						&collect_data($$node_ref[0], $$node_ref[1]);
						splice(@{$node_ref}, 0, 2);
					}
				}
			}
			else {
				$#{$node_ref} = -1;
			}
			if ( exists $$attr_ref{'table:number-columns-repeated'} ) {
				for (2..$$attr_ref{'table:number-columns-repeated'}) {
					$col++;
					$workbook{$table}[$row][$col] = $workbook{$table}[$row][$col - 1];
				}
			}
			if ( $text_p_count ) {		# only if IncludeCoveredCells is set
				$max_datarow = $row;
				if ( $col > $max_datacol ) { $max_datacol = $col; }
			}
		},
		'table:table-row'		=> sub	{
			if (( exists $$attr_ref{'table:visibility'} ) and ( $options{'DropHiddenRows'} )) {
				$#{$node_ref} = -1;
			}
			else {
				$row++;
				$workbook{$table}[$row] = undef;
				$col = -1;
				while ( @{$node_ref} ) {
					&collect_data($$node_ref[0], $$node_ref[1]);
					splice(@{$node_ref}, 0, 2);
				}
				if ( exists $$attr_ref{'table:number-rows-repeated'} ) {
					for (2..$$attr_ref{'table:number-rows-repeated'}) {
						$row++;
						$workbook{$table}[$row] = $workbook{$table}[$row - 1];	# copy reference, not data
					}
					if ( grep { defined $_ } @{$workbook{$table}[$row]} ) {
						$max_datarow = $row;
					}
				}
			}
		},
		'table:table-column'		=> sub {
			$col_count++;
			if ( $options{'DropHiddenColumns'} ) {
				if ( exists $$attr_ref{'table:visibility'} ) {
					$col_visible[$col_count] = 0;
				}
				else {
					$col_visible[$col_count] = 1;
				}
			}
			if ( exists $$attr_ref{'table:number-columns-repeated'} ) {
				for (2..$$attr_ref{'table:number-columns-repeated'} ) {
					$col_count++;
					$col_visible[$col_count] = $col_visible[$col_count - 1];
				}
			}
		},
		'table:table'			=> sub	{
			$table = $$attr_ref{'table:name'};
			$workbook{$table} = undef;
			$col_count = -1;
			$#col_visible = -1;
			$row = -1;
			$max_datarow = -1;
			$max_datacol = -1;
			while ( @{$node_ref} ) {
				&collect_data($$node_ref[0], $$node_ref[1]);
				splice(@{$node_ref}, 0, 2);
			}
			if ( ! $options{'NoTruncate'} ) {
				$#{$workbook{$table}} = $max_datarow;
				foreach ( @{$workbook{$table}} ) {
					$#{$_} = $max_datacol;
				}
			}
			if ( $options{'DropHiddenColumns'} ) {
				unless ( $#{$workbook{$table}} == -1 ) {	# Don't process empty tables
					my $width = $#{$workbook{$table}[0]};
					foreach ( @{$workbook{$table}} ) {
# Don't splice the row if it is a reference to the previous row (which has already been processed)
						unless ( $#{$_} < $width ) {
							for ( my $col = $#{$_}; $col >= 0; $col-- ) {
								if ( ! $col_visible[$col] )  {
									splice ( @{$_}, $col, 1 );
								}
							}
						}
					}
				}
			}
		},
	);
	if ( $begin_code{$tag} ) {
		$begin_code{$tag}->();
	}
	while ( @{$node_ref} ) {
		if ( $$node_ref[0] ) {		# Don't collect data if not in table nodes
			&collect_data($$node_ref[0], $$node_ref[1]);
		}
		splice(@{$node_ref}, 0, 2);
	}
}


1;
__END__
=head1 NAME

Spreadsheet::ReadSXC - Extract OpenOffice 1.x spreadsheet data


=head1 SYNOPSIS


  use Spreadsheet::ReadSXC qw(read_sxc);
  my $workbook_ref = read_sxc("/path/to/file.sxc");


  # Alternatively, unpack the .sxc file yourself and pass content.xml

  use Spreadsheet::ReadSXC qw(read_xml_file);
  my $workbook_ref = read_xml_file("/path/to/content.xml");


  # Alternatively, pass the XML string directly

  use Spreadsheet::ReadSXC qw(read_xml_string);
  use Archive::Zip;
  my $zip = Archive::Zip->new("/path/to/file.sxc");
  my $content = $zip->contents('content.xml');
  my $workbook_ref = read_xml_string($content);


  # Control the output through a hash of options (below are the defaults):

  my %options = (
	'ReplaceNewlineWith'	=> "",
	'IncludeCoveredCells'	=> 0,
	'DropHiddenRows'	=> 0,
	'DropHiddenColumns'	=> 0,
	'NoTruncate'		=> 0,
	'StandardDate'		=> 0,
  );
  my $workbook_ref = read_sxc("/path/to/file.sxc", \%options );


  # Iterate over every worksheet, row, and cell:

  use Unicode::String qw(utf8);

  foreach ( sort keys %$workbook_ref ) {
     print "Worksheet ", $_, " contains ", $#{$$workbook_ref{$_}} + 1, " row(s):\n";
     foreach ( @{$$workbook_ref{$_}} ) {
        foreach ( map { defined $_ ? $_ : '' } @{$_} ) {
	   print utf8(" '$_'")->as_string;
        }
        print "\n";
     }
  }


  # Cell D2 of worksheet "Sheet1"

  $cell = $$workbook_ref{"Sheet1"}[1][3];


  # Row 1 of worksheet "Sheet1":

  @row = @{$$workbook_ref{"Sheet1"}[0]};


  # Worksheet "Sheet1":

  @sheet = @{$$workbook_ref{"Sheet1"}};



=head1 DESCRIPTION


Spreadsheet::ReadSXC extracts data from OpenOffice 1.x spreadsheet
files (.sxc). It exports the function read_sxc() which takes a
filename and an optional reference to a hash of options as
arguments and returns a reference to a hash of references to
two-dimensional arrays. The hash keys correspond to the names of
worksheets in the OpenOffice workbook. The two-dimensional arrays
correspond to rows and cells in the respective spreadsheets.

If you prefer to unpack the .sxc file yourself, you can use the
function read_xml_file() instead and pass the path to content.xml
as an argument. Or you can extract the XML string from content.xml
and pass the string to the function read_xml_string(). Both
functions also take a reference to a hash of options as an
optional second argument.

Spreadsheet::ReadSXC requires XML::Parser to parse the XML
contained in .sxc files. It recursively traverses an XML tree to
find spreadsheet cells and collect their data. Only the contents
of text:p elements are returned, not the actual values of
table:value attributes. For example, a cell might have a
table:value-type attribute of "currency", a table:value attribute
of "-1500.99" and a table:currency attribute of "USD". The text:p
element would contain "-$1,500.99". This is the string which is
returned by the read_sxc() function, not the value of -1500.99.

Spreadsheet::ReadSXC was written with data import into an SQL
database in mind. Therefore empty spreadsheet cells correspond to
undef values in array rows. The example code above shows how to
replace undef values with empty strings.

If the .sxc file contains an empty spreadsheet its hash element will
point to an empty array (unless you use the 'NoTruncate' option in
which case it will point to an array of an array containing one
undefined element).

OpenOffice uses UTF-8 encoding. It depends on your environment how
the data returned by the XML Parser is best handled:

  use Unicode::String qw(latin1 utf8);
  $unicode_string = utf8($$workbook_ref{"Sheet1"}[0][0])->as_string;

  # this will not work for characters outside ISO-8859-1:

  $latin1_string = utf8($$workbook_ref{"Sheet1"}[0][0])->latin1;

Of course there are other modules than Unicode::String on CPAN that
handle conversion between encodings. It's your choice.

Table rows in .sxc files may have a "table:number-rows-repeated"
attribute, which is often used for consecutive empty rows. When you
format whole rows and/or columns in OpenOffice, it sets the numbers
of rows in a worksheet to 32,000 and the number of columns to 256, even
if only a few lower-numbered rows and cells actually contain data.
Spreadsheet::ReadSXC truncates such sheets so that there are no empty
rows after the last row containing data and no empty columns after the
last column containing data (unless you use the 'NoTruncate' option).

Still it is perfectly legal for an .sxc file to apply the
"table:number-rows-repeated" attribute to rows that actually contain
data (although I have only been able to produce such files manually,
not through OpenOffice itself). To save on memory usage in these cases,
Spreadsheet::ReadSXC does not copy rows by value, but by reference
(remember that multi-dimensional arrays in Perl are really arrays of
references to arrays). Therefore, if you change a value in one row, it
is possible that you find the corresponding value in the next row
changed, too:

  $$workbook_ref{"Sheet1"}[0][0] = 'new string';
  print $$workbook_ref{"Sheet1"}[1][0];

Keep in mind that after parsing a new .sxc file any reference previously
returned by the read_sxc() function will point to the new data structure.
Derefence the hash to save your data before parsing a new file. Or
derefence when calling the read_sxc() function:

  %workbook = %{$workbook_ref};
  %new_workbook = %{read_sxc("/path/to/newfile.sxc")};



=head1 OPTIONS

=over 4

=item ReplaceNewlineWith

By default, newlines within cells are ignored and all lines in a cell
are concatenated to a single string which does not contain a newline. To
keep the newline characters, use the following key/value pair in your
hash of options: 

  'ReplaceNewlineWith' => "\n"

However, you may replace newlines with any string you like.


=item IncludeCoveredCells

By default, the content of cells that are covered by other cells is
ignored because you wouldn't see it in OpenOffice unless you unmerge
the merged cells. To include covered cells in the data structure which
is returned by parse_sxc(), use the following key/value pair in your
hash of options:

  'IncludeCoveredCells' => 1


=item DropHiddenRows

By default, hidden rows are included in the data structure returned by
parse_sxc(). To drop those rows, use the following key/value pair in
your hash of options:

  'DropHiddenRows' => 1


=item DropHiddenColumns

By default, hidden columns are included in the data structure returned
by parse_sxc(). To drop those rows, use the following key/value pair
in your hash of options:

  'DropHiddenColumns' => 1


=item NoTruncate

By default, the two-dimensional arrays that contain the data within
each worksheet are truncated to get rid of empty rows below the last
row containing data and empty columns beyond the last column
containing data. If you prefer to keep those rows and columns, use the
following key/value pair in your hash of options:

  'NoTruncate' => 1


=item StandardDate

By default, date cells are returned as formatted. If you prefer to
obtain the date value as contained in the table:date-value attribute,
use the following key/value pair in your hash of options:

  'StandardDate' => 1

This option is a first step on the way to a different approach at
reading data from .sxc files. There should be more options to read in
values instead of the strings OpenOffice displays. It should give
more flexibility in working with the data obtained from OpenOffice
spreadsheets. 'float', 'percentage', and 'time' values should be
next. 'currency' is less obvious, though, as we need to consider
both its value and the 'table:currency' attribute. Formulas and
array formulas are yet another issue.


=back



=head1 SEE ALSO


http://books.evc-cit.info/book.html has extensive documentation of the
OpenOffice 1.x XML file format.



=head1 AUTHOR


Christoph Terhechte, E<lt>terhechte@cpan.orgE<gt>


=head1 COPYRIGHT AND LICENSE


Copyright 2005 by Christoph Terhechte

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself. 

=cut
