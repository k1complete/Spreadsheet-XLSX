package Spreadsheet::XLSX;

use 5.006000;
use strict;
use warnings;

use base 'Spreadsheet::ParseExcel::Workbook';

our $VERSION = '0.15';

use Archive::Zip;
use Spreadsheet::ParseExcel;
use Spreadsheet::XLSX::Fmt2007;

################################################################################

sub new {
    my ($class, $filename, $converter) = @_;

    my %shared_info;    # shared_strings, styles, style_info, rels, converter
    $shared_info{converter} = $converter;
    
    my $self = bless Spreadsheet::ParseExcel::Workbook->new(), $class;

    my $zip                     = __load_zip($filename);

    $shared_info{shared_strings} = __load_shared_strings($zip, $shared_info{converter});
    my ($styles, $style_info)   = __load_styles($zip);
    $shared_info{styles}        = $styles;
    $shared_info{style_info}    = $style_info;
    $shared_info{rels}          = __load_rels($zip);

    $self->_load_workbook($zip, \%shared_info);

    return $self;
}

sub _load_workbook {
    my ($self, $zip, $shared_info) = @_;

    my $member_workbook = $zip->memberNamed('xl/workbook.xml') or die("xl/workbook.xml not found in this zip\n");
    $self->{SheetCount} = 0;
    $self->{FmtClass}   = Spreadsheet::XLSX::Fmt2007->new;
    $self->{Flg1904}    = 0;
    if ($member_workbook->contents =~ /date1904="1"/) {
        $self->{Flg1904} = 1;
    }

    foreach ($member_workbook->contents =~ /\<(.*?)\/?\>/g) {

        /^(\w+)\s+/;

        my ($tag, $other) = ($1, $', $');

        my @pairs = split /\" /, $other;

        $tag eq 'sheet' or next;

        my $sheet = {
            MaxRow => 0,
            MaxCol => 0,
            MinRow => 1000000,
            MinCol => 1000000,
        };

        foreach ($other =~ /(\S+=".*?")/gsm) {

            my ($k, $v) = split /=?"/;    #"

            if ($k eq 'name') {
                $sheet->{Name} = $v;
                $sheet->{Name} = $shared_info->{converter}->convert($sheet->{Name}) if defined $shared_info->{converter};
            } elsif ($k eq 'r:id') {

                $sheet->{path} = $shared_info->{rels}->{$v};

            }

        }

        my $wsheet = Spreadsheet::ParseExcel::Worksheet->new(%$sheet);
        $self->{Worksheet}[$self->{SheetCount}] = $wsheet;
        $self->{SheetCount} += 1;

    }


    foreach my $sheet (@{$self->{Worksheet}}) {

        my $member_sheet = $zip->memberNamed("xl/$sheet->{path}") or next;

        my ($row, $col);

        my $parsing_v_tag = 0;
        my $s    = 0;
        my $s2   = 0;
        my $sty  = 0;
        foreach ($member_sheet->contents =~ /(\<.*?\/?\>|.*?(?=\<))/g) {
            my $rph = undef;
            my $ppr = undef;
            if (/^\<c\s*.*?\s*r=\"([A-Z])([A-Z]?)(\d+)\"/) {

                ($row, $col) = __decode_cell_name($1, $2, $3);

                $s   = m/t=\"s\"/      ? 1  : 0;
                $s2  = m/t=\"str\"/    ? 1  : 0;
                $sty = m/s="([0-9]+)"/ ? $1 : 0;

            } elsif (/^<v>/) {
                $parsing_v_tag = 1;
            } elsif (/^<\/v>/) {
                $parsing_v_tag = 0;
            } elsif (length($_) && $parsing_v_tag) {
                my $si = $shared_info->{shared_strings}->[$_];
                my $v = $s ? $si->{Val} : $_;
                if (exists($si->{Rph})) {
                    $rph = $s ? $si->{Rph} : undef;
                }
                if (exists($si->{PhoneticPr})) {
                    $ppr = $s ? $si->{PhoneticPr} : undef;
                }
                if ($v eq "</c>") {
                    $v = "";
                }
                my $type      = "Text";
                my $thisstyle = "";

                if (not($s) && not($s2)) {
                    $type = "Numeric";

                    if (defined $sty && defined $shared_info->{styles}->[$sty]) {
                        $thisstyle = $shared_info->{style_info}->{$shared_info->{styles}->[$sty]};
                        if ($thisstyle =~ /\b(mmm|m|d|yy|h|hh|mm|ss)\b/) {
                            $type = "Date";
                        }
                    }
                }


                $sheet->{MaxRow} = $row if $sheet->{MaxRow} < $row;
                $sheet->{MaxCol} = $col if $sheet->{MaxCol} < $col;
                $sheet->{MinRow} = $row if $sheet->{MinRow} > $row;
                $sheet->{MinCol} = $col if $sheet->{MinCol} > $col;

                if ($v =~ /(.*)E\-(.*)/gsm && $type eq "Numeric") {
                    $v = $1 / (10**$2);    # this handles scientific notation for very small numbers
                }

                my $cell = Spreadsheet::ParseExcel::Cell->new(
                    Val    => $v,
                    Format => $thisstyle,
                    Type   => $type
                );
                if ($s && $rph) {
                    $cell->{Rph} = $rph;
                    $cell->{PhoneticPr} = $ppr;
                }
                $cell->{_Value} = $self->{FmtClass}->ValFmt($cell, $self);
                if ($type eq "Date") {
                    if ($v < 1) {    #then this is Excel time field
                        $cell->{Type} = "Text";
                    }
                    $cell->{Val}  = $cell->{_Value};
                }
                $sheet->{Cells}[$row][$col] = $cell;
            }
        }

        $sheet->{MinRow} = 0 if $sheet->{MinRow} > $sheet->{MaxRow};
        $sheet->{MinCol} = 0 if $sheet->{MinCol} > $sheet->{MaxCol};

    }

    return $self;
}

# Convert cell name in the format AA1 to a row and column number.

sub __decode_cell_name {
    my ($letter1, $letter2, $digits) = @_;

    my $col = ord($letter1) - 65;

    if ($letter2) {
        $col++;
        $col *= 26;
        $col += (ord($letter2) - 65);
    }

    my $row = $digits - 1;

    return ($row, $col);
}

sub __load_rph_do {
    my ($si, $converter, @ret) = @_;
    if ($si =~ /<rPh(.*?)>(.*?)<\/rPh>(.*)/sm) {
        my $att = $1;
        my $cont = $2;
        my $nsi = $3;
        my $sb;
        my $eb;
        my $str;
        if ($att =~ /\s*sb\s*=\s*"(.*?)"/sm) {
            $sb = $1;
        }
        if ($att =~ /\s*eb\s*=\s*"(.*?)"/sm) {
            $eb = $1;
        }
        ## must be one time only(A.2 Spreadsheet ML line 1816)
        foreach my $t ($cont =~ /<t.*?>(.*?)<\/t/gsm) {
            $t = $converter->convert($t) if defined $converter;
            $str .= $t;
        }
        push @ret, {Val => $str, Sb => $sb, Eb => $eb};
        __load_rph_do($nsi, $converter, @ret);
    } else {
        return @ret;
    }
}
sub __load_rph {
    my ($si, $converter) = @_;
    my @ret = __load_rph_do($si, $converter, ());
    return \@ret;
}
sub __load_phonetic_pr {
    my ($si) = @_;
    my %retval = ();
    foreach my $i ($si =~ /<phoneticPr\s(.*?)\/>/gsm) {
        if ($i =~ /type\s*=\s*"(.*?)"/) {
            $retval{Type} = $1;
        }
        if ($i =~ /fontId\s*=\s*"(.*?)"/) {
            $retval{FontId} = $1;
        }
    }
    return \%retval;
}
sub __load_shared_strings {
    my ($zip, $converter) = @_;

    my $member_shared_strings = $zip->memberNamed('xl/sharedStrings.xml');

    my @shared_strings = ();

    if ($member_shared_strings) {
        my $mstr = $member_shared_strings->contents;
        $mstr =~ s/<t\/>/<t><\/t>/gsm;    # this handles an empty t tag in the xml <t/>
        foreach my $si ($mstr =~ /<si.*?>(.*?)<\/si/gsm) {
            
            my $rph = __load_rph($si, $converter);
            my $phoneticPr = __load_phonetic_pr($si);
            my $str;
            $si =~ s/<rPh.*?<\/rPh>//gsm;
            foreach my $t ($si =~ /<t.*?>(.*?)<\/t/gsm) {
                $t = $converter->convert($t) if defined $converter;
                $str .= $t;
            }
            push @shared_strings, {Val => $str, Rph => $rph, 
                                                PhoneticPr => $phoneticPr};
        }
    }

    return \@shared_strings;
}


sub __load_styles {
    my ($zip) = @_;

    my $member_styles = $zip->memberNamed('xl/styles.xml');

    my @styles = ();
    my %style_info = ();

    if ($member_styles) {
        my $formatter = Spreadsheet::XLSX::Fmt2007->new();

        foreach my $t ($member_styles->contents =~ /xf\ numFmtId="([^"]*)"(?!.*\/cellStyleXfs)/gsm) {    #"
            push @styles, $t;
        }

        my $default = $1 || '';
    
        foreach my $t1 (@styles) {
            $member_styles->contents =~ /numFmtId="$t1" formatCode="([^"]*)/;
            my $formatCode = $1 || '';
            if ($formatCode eq $default || not($formatCode)) {
                if ($t1 == 9 || $t1 == 10) {
                    $formatCode = '0.00000%';
                } elsif ($t1 == 14) {
                    $formatCode = 'yyyy-mm-dd';
                } else {
                    $formatCode = '';
                }
#                $formatCode = $formatter->FmtStringDef($t1);
            }
            $style_info{$t1} = $formatCode;
            $default = $1 || '';
        }

    }
    return (\@styles, \%style_info);
}


sub __load_rels {
    my ($zip) = @_;

    my $member_rels = $zip->memberNamed('xl/_rels/workbook.xml.rels') or die("xl/_rels/workbook.xml.rels not found in this zip\n");

    my %rels = ();

    foreach ($member_rels->contents =~ /\<Relationship (.*?)\/?\>/g) {

        my ($id, $target);
        ($id) = /Id="(.*?)"/;
        ($target) = /Target="(.*?)"/;
 
	    if (defined $id and defined $target) {	
    		$rels{$id} = $target;
        }

    }

    return \%rels;
}

sub __load_zip {
    my ($filename) = @_;

    my $zip = Archive::Zip->new();

    if (ref $filename) {
        $zip->readFromFileHandle($filename) == Archive::Zip::AZ_OK or die("Cannot open data as Zip archive");
    } else {
        $zip->read($filename) == Archive::Zip::AZ_OK or die("Cannot open $filename as Zip archive");
    }
    
    return $zip;
}


1;
__END__
=encoding utf8
=head1 NAME

Spreadsheet::XLSX - Perl extension for reading MS Excel 2007 files;

=head1 SYNOPSIS

 use Text::Iconv;
 my $converter = Text::Iconv -> new ("utf-8", "windows-1251");
 
 # Text::Iconv is not really required.
 # This can be any object with the convert method. Or nothing.

 use Spreadsheet::XLSX;
 
 my $excel = Spreadsheet::XLSX -> new ('test.xlsx', $converter);
 
 foreach my $sheet (@{$excel -> {Worksheet}}) {
 
 	printf("Sheet: %s\n", $sheet->{Name});
 	
 	$sheet -> {MaxRow} ||= $sheet -> {MinRow};
 	
         foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
         
 		$sheet -> {MaxCol} ||= $sheet -> {MinCol};
 		
 		foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
 		
 			my $cell = $sheet -> {Cells} [$row] [$col];
 
 			if ($cell) {
 			    printf("( %s , %s ) => %s\n", $row, $col, $cell -> {Val});
 			}
 
 		}
 
 	}
 
 }

=head1 DESCRIPTION

This module is a (quick and dirty) emulation of Spreadsheet::ParseExcel for 
Excel 2007 (.xlsx) file format.  It supports styles and many of Excel's quirks, 
but not all.  It populates the classes from Spreadsheet::ParseExcel for interoperability; 
including Workbook, Worksheet, and Cell.

=head2 Phonetic hint

Phonetic hint, used in far east asia is supported by 'Rph' and 'PhoneticPr'
key:
    <si>
     <t>課きく 毛こ</t>
     <rPh sb="0" eb="1">
      <t>カ</t> 
     </rPh>
     <rPh sb="4" eb="5"> 
      <t>ケ</t>
     </rPh>
     <phoneticPr fontId="1"/>
    </si>

if a cell[0][0]->{Val} is '課きく毛こ', The
cell[0][0]->{Rph}->[0]->{Val} is 'カ', cell[0][0]->{Rph}->[0]->{Sb} is
0, and cell[0][0]->{Rph}->[0]->{Eb} is 1,
cell[0][0]->{Rph}->[1]->{Val} is 'ケ', cell[0][0]->{Rph}->[1]->{Sb} is
4, and cell[0][0]->{Rph}->[0]->{Eb} is 5,
cell[0][0]->{PhoneticPr}->{FontId} is 1,
cell[0][0]->{PhoneticPr}->{Type} is undef.  Phonetic hint keys are
named by capitalizing the first letter from ecma-376 attribute names.

Sb is base text start index, Eb is base text end index(cf. ecma-376
18.4.6).  if {PhoneticPr}->{Type} is undef, then "fullWithKatakana" by
default (cf. ecma-276 B.2 Spreadsheet ML line 2005).

=head1 SEE ALSO

=over 2

=item Text::CSV_XS, Text::CSV_PP

http://search.cpan.org/~hmbrand/

A pure perl version is available on http://search.cpan.org/~makamaka/

=item Spreadsheet::ParseExcel

http://search.cpan.org/~kwitknr/

=item Spreadsheet::ReadSXC

http://search.cpan.org/~terhechte/

=item Spreadsheet::BasicRead

http://search.cpan.org/~gng/ for xlscat likewise functionality (Excel only)

=item Spreadsheet::ConvertAA

http://search.cpan.org/~nkh/ for an alternative set of cell2cr () /
cr2cell () pair

=item Spreadsheet::Perl

http://search.cpan.org/~nkh/ offers a Pure Perl implementation of a
spreadsheet engine. Users that want this format to be supported in
Spreadsheet::Read are hereby motivated to offer patches. It's not high
on my todo-list.

=item xls2csv

http://search.cpan.org/~ken/ offers an alternative for my C<xlscat -c>,
in the xls2csv tool, but this tool focusses on character encoding
transparency, and requires some other modules.

=item Spreadsheet::Read

http://search.cpan.org/~hmbrand/ read the data from a spreadsheet (interface 
module)

=back

=head1 AUTHOR

Dmitry Ovsyanko, E<lt>do@eludia.ru<gt>, http://eludia.ru/wiki/

Patches by:

	Steve Simms
	Joerg Meltzer
	Loreyna Yeung	
	Rob Polocz
	Gregor Herrmann
	H.Merijn Brand
	endacoe
	Pat Mariani
	Sergey Pushkin
	
=head1 ACKNOWLEDGEMENTS	

	Thanks to TrackVia Inc. (http://www.trackvia.com) for paying for Rob Polocz working time.

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2008 by Dmitry Ovsyanko

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself, either Perl version 5.8.8 or,
at your option, any later version of Perl 5 you may have available.

cf. http://www.ecma-international.org/publications/standards/Ecma-376.htm

=cut
