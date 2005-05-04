package Spreadsheet::ReadSXC;

use 5.006;
use strict;
use warnings;

require Exporter;

our @ISA = qw(Exporter);
our @EXPORT_OK = qw(read_sxc);
our $VERSION = '0.03';

use Archive::Zip;
use XML::Parser::Lite::Tree;

my %workbook;
my $table;
my $row;
my $cell;
my $max_datarow;
my $max_datacol;

sub read_sxc {
	my $zip = Archive::Zip->new(shift);
	my $content = $zip->contents('content.xml');
	my $tree_parser = XML::Parser::Lite::Tree::instance();
	my $tree = $tree_parser->parse($content);
	%workbook = ();
	&collect_data($tree);
	return \%workbook;
}

sub collect_data {
	my $node_ref = shift;
	if ( $node_ref->{type} eq 'tag' ) {
		if ( $node_ref->{name} eq 'table:table' ) {
			$table = $node_ref->{attributes}->{'table:name'};
			$workbook{$table} = undef;
			$row = -1;
			$max_datarow = -1;
			$max_datacol = -1;
			for my $i (0..$#{$node_ref->{children}}) {
				&collect_data($node_ref->{children}[$i]);
			}
			$#{$workbook{$table}} = $max_datarow;
			foreach ( @{$workbook{$table}} ) {
				$#{$_} = $max_datacol;
			}
		}
		elsif ( $node_ref->{name} eq 'table:table-row' ) {
			$row++;
			$workbook{$table}[$row] = undef;
			$cell = -1;
			for my $i (0..$#{$node_ref->{children}}) {
				&collect_data($node_ref->{children}[$i]);
			}
			if ( exists $node_ref->{attributes}->{'table:number-rows-repeated'} ) {
				for (2..$node_ref->{attributes}->{'table:number-rows-repeated'}) {
					$row++;
					$workbook{$table}[$row] = $workbook{$table}[$row-1];
				}
			}
		}
		elsif ( $node_ref->{name} eq 'table:table-cell' ) {
			$cell++;
			$workbook{$table}[$row][$cell] = undef;
			for my $i (0..$#{$node_ref->{children}}) {
				&collect_data($node_ref->{children}[$i]);
			}
			if ( exists $node_ref->{attributes}->{'table:number-columns-repeated'} ) {
				for (2..$node_ref->{attributes}->{'table:number-columns-repeated'}) {
					$cell++;
					$workbook{$table}[$row][$cell] = $workbook{$table}[$row][$cell - 1];
				}
			}
		}
		else {
			for my $i (0..$#{$node_ref->{children}}) {
				&collect_data($node_ref->{children}[$i]);
			}
		}
	}
	elsif ( $node_ref->{type} eq 'data' ) {
		if ( defined $table ) {
			$workbook{$table}[$row][$cell] .= $node_ref->{content};
			$max_datarow = $row;
			if ( $cell > $max_datacol ) {
				$max_datacol = $cell;
			}
		}
	}
	elsif ( $node_ref->{type} eq 'root') {
		for my $i (0..$#{$node_ref->{children}}) {
			&collect_data($node_ref->{children}[$i]);
		}
	}
	else {
		die "This doesn't look like an OpenOffice Workbook.\n"
	}
}

1;
__END__
=head1 NAME

Spreadsheet::ReadSXC - Extract OpenOffice 1.x spreadsheet data


=head1 SYNOPSIS


  use Spreadsheet::ReadSXC qw(read_sxc);
  my $workbook_ref = read_sxc("/path/to/file.sxc");


Iterate over every worksheet, row, and cell:


  foreach ( sort keys %$workbook_ref ) {
     print "Worksheet ", $_, " contains ", $#{$$workbook_ref{$_}} + 1, " row(s):\n";
     foreach ( @{$$workbook_ref{$_}} ) {
        foreach ( map { defined $_ ? $_ : '' } @{$_} ) {
	   print " '$_'";
        }
        print "\n";
     }
  }


Cell D2 of worksheet "Sheet1"


  $cell = $$workbook_ref{"Sheet1"}[1][3];


Decode the contained XML to latin1


  use Unicode::String qw(latin1 utf8);
  use XML::Quote;
  $cell_text = utf8(xml_dequote($$workbook_ref{"Sheet1"}[1][3]))->latin1;


Row 1 of worksheet "Sheet1":


  @row = @{$$workbook_ref{"Sheet1"}[0]};


Worksheet "Sheet1":


  @sheet = @{$$workbook_ref{"Sheet1"}};



=head1 DESCRIPTION


This is a very simple module for extracting data from OpenOffice 1.x
spreadsheet files (.sxc). It exports only one function read_sxc()
which takes a filename as an argument and returns a hash of references
to two-dimensional arrays. The hash keys correspond to the names of
worksheets in the OpenOffice workbook. The two-dimensional arrays
correspond to rows and cells in the respective spreadsheets.

Spreadsheet::ReadSXC requires XML::Parser::Lite::Tree to parse the
XML contained in .sxc files. It recursively traverses the XML tree
to find spreadsheet cells and collect their data. Only the contents
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
point to an empty array.

Table rows in .sxc files may have a "table:number-rows-repeated"
attribute, which is often used for consecutive empty rows. When you
format whole rows and/or columns in OpenOffice, it sets the numbers
of rows in a worksheet to 32,000 and the number of columns to 256, even
if only a few lower-numbered rows and cells actually contain data.
Spreadsheet::ReadSXC truncates such sheets so that there are no empty
rows after the last row containing data and no empty columns after the
last column containing data.

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
