use Win32::OLE qw(in);
use JIRA::REST;
use JIRA::Client::Automated;
use Data::Dumper;
use JSON;
use utf8;
use Encode qw(encode_utf8);
my $Outlook = Win32::OLE->GetObject("Outlook.Application")
 || Win32::OLE->new("Outlook.Application");

my $Session = $Outlook->Session;
#my $NameSpace = $Session->GetNameSpace("MAPI");
#my $ContactsFolder = $NameSpace->GetDefaultFolder(olFolderContacts);
$Patch_Name="R_2.4.1_68355_non_CS_dotNET_SSRS_Report_Petredec_SP196";
$bool=1;
ShowSubFolders ($Session->Folders);
#################################################
sub ShowSubFolders{
my $FolderCollection = shift;
  my $level = shift || 0;
#print "HI";
  for my $Folder (in $FolderCollection) {
    #print "  " x $level . "Folder$level: $Folder->{Name}\t(\n";# . $Folder->Items->{Count}. ")\n";
    #if ($Folder->{Name} =~ /Personal Folders/) {
    if ($Folder->{Name} =~ /Goutham, Duggi/) {
	#print "hi";
      ShowSubFolders($Folder->Folders, $level + 1);
    }
	if ($Folder->{Name} =~ /Inbox/) {
	#print "hi";
      ShowSubFolders($Folder->Folders, $level + 1);
    }
    if ($Folder->{Name} =~ /to_do/) { 
      my $msgCount=1;
      print "*** Please Allow the msgbox in outlook that requests external program access**\n";
      #$Folder->Items->Sort ("CreationTime",1); # Descending by date
      for my $Item (in $Folder->Items) {
	  my $line=$Item->{Body};
	  # print "$line\n";
	  @subjects=$Item->{subject};
	  foreach $subject(@subjects)
	  {
		print "Subject is : ".$subject."\n";
		my $Body=$Item->{Body};
		@lines=split(/\n/,$Body);
		for my $line(0 .. $#lines)
		{
			if($bool)
			{
			if($lines[$line]=~/^Changeset:/)
			{
		  #print $line."\n";
				$changeset=substr($lines[$line],length("Changeset:"),length($lines[$line]));
		  #print "Changeset is ".$changeset."\n";
			}
			if($lines[$line]=~/Comment:/)
			{
				$comment=$lines[$line+1];
		   #print "Comment is  :".$comment."\n";
				@issue=$comment=~/I#(\w+\-\d{0,})/ig;
				@issue=&number(@issue);
			}
			if($lines[$line]=~/Items:/)
			{
				#print "Changeset is $changeset \n";
				#print "Issue length is @issue \n";
				push @Aoh,{changesets=>"$changeset",issue=>"@issue"};
				$bool=0;
			}
		 }
		 $bool=1;
	  }
	  }
	  for $i(0 .. $#Aoh)
	  {
	     
	     #print " is $Aoh[$i]{issue} \n";
		 if($Aoh[$i]{changesets} ne "")
		 {
		    $changeset=trim($Aoh[$i]{changesets});
			@issues=split(/ /,$Aoh[$i]{issue});
			foreach(@issues)
			{
			  if($_ ne "")
			  {
			  &connecttoJira($changeset,$_,$Patch_Name);
				}
			}
		 }
	  }
	  
	  # if(index($line,"Patch Name")!=-1){
	  # my $value=substr($line,index($line,"Patch Name"));
	  # print $value;
	   	   	  }
			  }
	  }	  
	  }
sub number(){
	(@number) = @_;
	chomp(@number);
	@final_num = ();
	foreach(@number){
	    $_ =~ s/\n\r\t\s//g;
		next if($_ eq "");
		@num = $_=~/(\w+\-\d{0,})|(\d{0,})/g;
		# print "function @num\n";
		push(@final_num,@num);
	}	
	return @final_num;
}

sub uniq {
    return keys %{{ map { $_ => 1 } @_ }};
}
sub trim($)
{
	my $string = shift;
	$string =~ s/^\s+//;
	$string =~ s/\s+$//;
	return $string;
}
sub connecttoJira()
{
my $changeset=shift;
my $issue=shift;
my $patchname=shift;
 
print "Connecting to JIra for $issue.....\n";
print "PatchName is  $patchname.....\n";
print "Changeset is  $changeset.....\n";
 

# my $url='http://172.16.170.182:1234/';
# #my $key='ADSO-2';
# my $const='/issue/';
# my $jira=JIRA::Client::Automated->new($url,'gduggi','India123!');
# #my $issue=$jira->GET($const.$key);
# my $comment1="Patched in $patchname and changeset is : $changeset";
# $jira->create_comment($issue, $comment1);
# my $operation="In Review";
# #$jira->close_issue($key,$resolve,$comment,$update_hash,$operation);
# print "Connecting to JIra for $issue.....\n";
# my $content = $jira->get_issue($issue);
# $value=to_json($content);
# my $encode=encode_utf8( $value);
# my $text=decode_json($encode);
# $status=$text->{'fields'}{'status'}{'name'};
# print "$issue : Satus is : ".$status."\n";
# if($status=="In Progress")
# {
   # $jira->close_issue($issue,$resolve,$comment1,$update_hash,$operation);
# }
}	
