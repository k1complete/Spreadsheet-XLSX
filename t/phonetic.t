use Test::More tests => 13;

      use Spreadsheet::XLSX;
      use warnings;
      
      my $fn = __FILE__;
      $fn =~ s{t$}{xlsx};

      my $excel = Spreadsheet::XLSX->new($fn);
      my $cells = $excel->{Worksheet}[0]{Cells};
      ok ($cells->[0][0]->value() eq '佐藤', 'shared text');
      ok ($cells->[0][0]->{Val} eq '佐藤', 'shared text');
      ok ($cells->[0][0]->{RPh}->[0]->{Val} eq 'サトウ', 'phonetic');
      ok ($cells->[5][0]->{Val} eq '十分', 'phonetic');
      ok ($cells->[6][0]->{Val} eq '課きく 毛こ', 'kakiku keko');
      ok ($cells->[6][0]->{RPh}->[0]->{Val} eq 'カ', 'phonetic_ka');
      ok ($cells->[6][0]->{RPh}->[0]->{Sb} eq '0', 'phonetic_sb_ka');
      ok ($cells->[6][0]->{RPh}->[0]->{Eb} eq '1', 'phonetic_eb_ka');
      ok ($cells->[6][0]->{RPh}->[1]->{Val} eq 'ケ', 'phonetic_ke');
      ok ($cells->[6][0]->{RPh}->[1]->{Sb} eq '4', 'phonetic_sb_ke');
      ok ($cells->[6][0]->{RPh}->[1]->{Eb} eq '5', 'phonetic_eb_ke');
      ok ($cells->[6][0]->{PhoneticPr}->{Type} eq "fullwidthKatakana", 'phonetic_property type');
      ok ($cells->[6][0]->{PhoneticPr}->{FontId} eq '1', 'phonetic_property font');

 
