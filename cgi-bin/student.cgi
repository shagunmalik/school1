#!C:\strawberry\perl\bin\perl.exe
use Spreadsheet::ParseExcel;
use Spreadsheet::XLSX;
use strict;
use CGI;
use Data::Dumper;
print "Content-type: text/html; charset=iso-8859-1\n\n";

my $q = new CGI;

my $displayMessage= '';

if($q->param('submit_student_profile')){
	my $upfile = $q->param('upfile');
    my $basename = GetBasename($upfile);
    my @filearr=split(/\./,$basename);
    my $ext=$filearr[scalar(@filearr) - 1];
	
	my $dt=get_timestamp();
	my $path='C:\\wamp\\bin\\apache\\apache2.2.22\\cgi-bin\\school\\data\\';
	
	my $flname="student_profile_$dt.$ext";
	my $file=$path.$flname;
	if (! open(OUTFILE, ">$file") ) {
		$displayMessage =qq{Can't open excel for writing - $!};
	} else	{
		my $nBytes = 0;
		my $totBytes = 0;
		my $buffer = "";
		binmode($upfile);
		while ( $nBytes = read($upfile, $buffer, 1024) ) {
			print OUTFILE $buffer;
			$totBytes += $nBytes;
		}
		close(OUTFILE);
		if(-e $file)
		{
			my ($dataref,$emailsref)=process_file($file,$ext);
		}
	}
}

sub process_file
{
	my $filename=shift;
	my $ext=shift;
	my @data;
	$filename = 'C:\\wamp\\bin\\apache\\apache2.2.22\\cgi-bin\\school\\data\\student_profile.'.$ext;
	if($ext eq 'xlsx')
	{
		my $excel = Spreadsheet::XLSX->new($filename,);
		foreach my $sheet (@{$excel -> {Worksheet}}) {
			my @key_array=();
			
		    $sheet -> {MaxRow} ||= $sheet -> {MinRow};
			$sheet -> {MaxCol} ||= $sheet -> {MinCol};
			
            foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
				my $cell = $sheet->{Cells} [0] [$col];
				push @key_array ,  $cell->{Val} if($cell); 
			}
		    foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
		        if($row > 0) {
		            my $r;
					foreach my $index (0..$#key_array){
						$r->{$key_array[$index]} = $sheet->{Cells}[$row][$index]->{Val} ; 
					}
					push @data ,  $r;	          
		        }
		    }
		}
	} else {   
		my $oExcel = new Spreadsheet::ParseExcel;
		print "You must provide a filename to $filename to be parsed as an Excel file" unless $filename;
		my $oBook = $oExcel->Parse($filename) || print "Error in reading file";
		my($iR, $iC, $sheet, $oWkC);
#        print  "FILE  :", $oBook->{File} , "\n";
#        print  "COUNT :", $oBook->{SheetCount} , "\n";
#        print  "AUTHOR:", $oBook->{Author} , "\n" if defined $oBook->{Author};

		for(my $iSheet=0; $iSheet < $oBook->{SheetCount} ; $iSheet++)
		{
			my @key_array=();
			$sheet = $oBook->{Worksheet}[$iSheet];
			print "<br> MinRow = $sheet->{MinRow}, $sheet->{MaxRow} <br>";
			for(my $iC = $sheet->{MinCol} ; defined $sheet->{MaxCol} && $iC <= $sheet->{MaxCol} ;    $iC++){
			    $oWkC = $sheet->{Cells}[0][$iC];
			    push @key_array , $oWkC->Value if($oWkC);
			}
			print Dumper(\@key_array);
			for(my $row = $sheet->{MinRow} ;defined $sheet->{MaxRow} && $row <= $sheet->{MaxRow} ;$row++)
			{
				
				if($row > 0)
				{
					my $r;
					foreach my $index (0..$#key_array){
						$r->{$key_array[$index]} = $sheet->{Cells}[$row][$index]->{Val} ; 
					}
					push @data ,  $r;
				}
			}
		}
	}
	print "<pre>".Dumper(\@data)."</pre>";
	return \@data;
}

sub GetBasename {
    my $fullname = shift;

    my(@parts);
    # check which way our slashes go.
    if ( $fullname =~ /(\\)/ ) {
        @parts = split(/\\/, $fullname);
    } else {
        @parts = split(/\//, $fullname);
    }

    return(pop(@parts));
}

sub get_timestamp {
    my $typ=shift;
    my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
    $year += 1900;
    $mon++;
    if(!$typ){
        return $year . sprintf('%02u%02u%02u%02u%02u', $mon, $mday, $hour, $min, $sec);  
    }
    else{
        return $year . sprintf('-%02u-%02u', $mon, $mday);
    }
}

print $displayMessage;
print qq{
<html>
<head>

</head>
<body>
 <form name="student_profile" method="post" enctype="multipart/form-data">
  <div align="center">
   Upload excel : <input type="file" name="upfile" > <br>
   <input type="submit" value="Upload Student Profiles" name="submit_student_profile">
  </div>
 </form>
</body>
</html>
};

