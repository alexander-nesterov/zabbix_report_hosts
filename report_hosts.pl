use strict;
use warnings;

use Excel::Writer::XLSX;
use MIME::Lite;
use JSON::RPC::Client;
use Data::Dumper;
use utf8;

#================================================================
#Constants
#================================================================
#ZABBIX
use constant ZABBIX_USER			=> 'Admin';
use constant ZABBIX_PASSWORD	=> '78%Ytn#4Wq32!Hynm90';
use constant ZABBIX_SERVER		=> '109.120.174.24';

#DEBUG
use constant DEBUG		=> 0; #0 - False, 1 - True

#================================================================
##Global variables
#================================================================
my $ZABBIX_AUTH_ID;

binmode(STDOUT,':utf8');

main();

#================================================================
sub zabbix_auth
{
    my %data;

    $data{'jsonrpc'} = '2.0';
    $data{'method'} = 'user.login';
    $data{'params'}{'user'} = ZABBIX_USER;
    $data{'params'}{'password'} = ZABBIX_PASSWORD;
    $data{'id'} = 1;

    my $response = send_to_zabbix(\%data);

    if (!defined($response))
    {
				print "Authentication failed, zabbix server: " . ZABBIX_SERVER . "\n" if DEBUG;
				return 0;
    }

    $ZABBIX_AUTH_ID = $response->content->{'result'};

    if (!defined($ZABBIX_AUTH_ID)) 
    {
				print "Authentication failed, zabbix server: " . ZABBIX_SERVER . "\n" if DEBUG;
				return 0; 
    }

    print "Authentication successful. Auth ID: $ZABBIX_AUTH_ID\n" if DEBUG;

    undef $response;

    return 1;
}


#================================================================
sub send_to_zabbix
{
    my $json = shift;

    my $response;

    my $url = "http://" . ZABBIX_SERVER . "/api_jsonrpc.php";

    my $client = new JSON::RPC::Client;

    $response = $client->call($url, $json);

    return $response;
}

#================================================================
sub zabbix_logout
{
    my %data;

    $data{'jsonrpc'} = '2.0';
    $data{'method'} = 'user.logout';
    $data{'params'} = [];
    $data{'auth'} = $ZABBIX_AUTH_ID;
    $data{'id'} = 1;

    my $response = send_to_zabbix(\%data);

    if (!defined($response))
    {
				print "Logout failed, zabbix server: " . ZABBIX_SERVER . "\n" if DEBUG;
				return 0;
    }

    print "Logout successful. Auth ID: $ZABBIX_AUTH_ID\n" if DEBUG;

    undef $response;
}

#================================================================
sub get_hosts
{
	my %data;

    $data{'jsonrpc'} = '2.0';
    $data{'method'} = 'host.get';
    
    #https://www.zabbix.com/documentation/3.2/manual/api/reference/host/object
    my @hosts_params = ('hostid', 'host', 'name', 'description', 'available', 'status');
	$data{'params'}{'output'} = [@hosts_params];
		
	#https://www.zabbix.com/documentation/3.2/manual/api/reference/hostinterface/object
	my @interfaces_params = ('ip', 'dns', 'port');
	$data{'params'}{'selectInterfaces'} = [@interfaces_params];
    
    #https://www.zabbix.com/documentation/3.2/manual/api/reference/item/object
    my @items_params = ('itemid', 'name', 'type', 'key_', 'value_type', 'delay', 'description', 'status', 'history', 'trends');
    $data{'params'}{'selectItems'} = [@items_params];
    
    #https://www.zabbix.com/documentation/3.2/manual/api/reference/trigger/object
    my @triggers_params = ('triggerid', 'description', 'expression', 'comments', 'priority', 'status');
    $data{'params'}{'selectTriggers'} = [@triggers_params];
    
    #https://www.zabbix.com/documentation/3.2/manual/api/reference/template/object
    my @templates_params = ('templateid', 'host', 'description');
    $data{'params'}{'selectParentTemplates'} = [@templates_params];
        
    #$data{'params'}{'selectGroups'} = 'extend';
    
    $data{'params'}{'sortfield'} = 'name'; #Possible values are: hostid, host, name, status.
    $data{'params'}{'sortorder'} = 'DESC'; #DESC or ASC
    
    $data{'auth'} = $ZABBIX_AUTH_ID;
    $data{'id'} = 1;
    
    my $response = send_to_zabbix(\%data);
    
    #print Dumper $response;
    foreach my $host(@{$response->content->{'result'}}) 
    {
		foreach (@hosts_params)
		{
			print 'Host: ' . $host->{$_} . "\n";
		}
				
		#get_hosts_params('Host',\@hosts_params, \@{$response->content->{'result'}});
		get_hosts_params('Interface', \@interfaces_params, \@{$host->{'interfaces'}});
		get_hosts_params('Item', \@items_params, \@{$host->{'items'}});
		get_hosts_params('Trigger', \@triggers_params, \@{$host->{'triggers'}});
		get_hosts_params('Template', \@templates_params, \@{$host->{'templates'}});
	}
}

#================================================================
sub get_hosts_params
{
	my ($label, $ref_params, $ref_data) = @_;
		
	foreach my $data(@{$ref_data})
	{					
		foreach my $params(@{$ref_params})
		{   
			print "$label: " . $data->{$params} . "\n";
		}
	}
}
#================================================================
sub main
{
	if (zabbix_auth() != 0)
    {
		get_hosts();
    } 
    print "*** Done ***\n";
}