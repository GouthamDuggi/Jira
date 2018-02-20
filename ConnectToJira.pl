use JIRA::REST;
use JIRA::Client::Automated;
use Data::Dumper;
use JSON;
use utf8;
use Encode qw(encode_utf8);
my $url='http://172.16.170.182:1234/';
my $key='ADSO-3';
my $const='/issue/';
my $jira=JIRA::Client::Automated->new($url,'gduggi','India123!');
#my $issue=$jira->GET($const.$key);
#my $comment="Patched in So and so branch and changeset is : 7900";
#$jira->create_comment($key, $comment);
my $operation="In Review";
#$jira->close_issue($key,$resolve,$comment,$update_hash,$operation);
my $issue = $jira->get_issue($key);
$value=to_json($issue);
my $encode=encode_utf8( $value);
my $text=decode_json($encode);
$status=$text->{'fields'}{'status'}{'name'};
print "Satus is : ".$status."\n";
if($status=="In Progress")
{
   #$jira->close_issue($key,$resolve,$comment,$update_hash,$operation);
}



