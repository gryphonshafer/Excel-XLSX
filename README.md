# NAME

Excel::XLSX - Read and write Excel XLSX data

# VERSION

version 1.03

[![test](https://github.com/gryphonshafer/Excel-XLSX/workflows/test/badge.svg)](https://github.com/gryphonshafer/Excel-XLSX/actions?query=workflow%3Atest)
[![codecov](https://codecov.io/gh/gryphonshafer/Excel-XLSX/graph/badge.svg)](https://codecov.io/gh/gryphonshafer/Excel-XLSX)

# SYNOPSIS

    use Excel::XLSX qw( from_xlsx to_xlsx );

    open( my $in, '<:raw', 'in.xlsx' ) or die $!;
    local $/ = undef;
    my $xlsx_in = <$in>;

    my $workbook_data = from_xlsx($xlsx_in);

    $workbook_data->{worksheets}[0]{cells}{12}{7}{value} = 'Hello world!';

    my $xlsx = to_xlsx($workbook_data);

    open( my $out, '>:raw', 'out.xlsx' ) or die $!;
    print $out $xlsx;

# DESCRIPTION

This module offers for export 2 functions, `from_xlsx` and `to_xlsx`, that
will read from raw binary Excel XLSX data into a data structure
and write from a data structure to raw binary Excel XLSX data.

# FUNCTIONS

## from\_xlsx

Reads from raw binary Excel XLSX data and returns a data structure.

    my $workbook_data = from_xlsx($xlsx_in);

## to\_xlsx

Requires a data structure like what might be returned from `from_xlsx` and
returns raw binary Excel XLSX data.

    my $xlsx = to_xlsx($workbook_data);

# DATA STRUCTURE

The data structure is expected to generally look like:

    formats:
      - font: Arial
        size: 12
      - font: Arial
        size: 10
        color: #00FF00
    worksheets:
      - name: Example Worksheet
        cells:
            12:
                7:
                  - format_id: 0
                    value: Hello world!

# SEE ALSO

[Spreadsheet::ParseXLSX](https://metacpan.org/pod/Spreadsheet%3A%3AParseXLSX), [Excel::Writer::XLSX](https://metacpan.org/pod/Excel%3A%3AWriter%3A%3AXLSX), [Excel::ValueReader::XLSX](https://metacpan.org/pod/Excel%3A%3AValueReader%3A%3AXLSX),
[Excel::ValueWriter::XLSX](https://metacpan.org/pod/Excel%3A%3AValueWriter%3A%3AXLSX).

You can also look for additional information at:

- [GitHub](https://github.com/gryphonshafer/Excel-XLSX)
- [MetaCPAN](https://metacpan.org/pod/Excel::XLSX)
- [GitHub Actions](https://github.com/gryphonshafer/Excel-XLSX/actions)
- [Codecov](https://codecov.io/gh/gryphonshafer/Excel-XLSX)
- [CPANTS](http://cpants.cpanauthors.org/dist/Excel-XLSX)
- [CPAN Testers](http://www.cpantesters.org/distro/S/Excel-XLSX.html)

# AUTHOR

Gryphon Shafer <gryphon@cpan.org>

# COPYRIGHT AND LICENSE

This software is Copyright (c) 2025-2050 by Gryphon Shafer.

This is free software, licensed under:

    The Artistic License 2.0 (GPL Compatible)
