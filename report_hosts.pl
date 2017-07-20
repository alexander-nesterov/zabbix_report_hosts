#!/usr/bin/perl -w

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
use constant ZABBIX_USER     => 'Admin';
use constant ZABBIX_PASSWORD => 'password';
use constant ZABBIX_SERVER   => 'localhost';

#DEBUG
use constant DEBUG => 1; #0 - False, 1 - True

my %TEST_STATUS = (
    0 => {
	color => '#00AA00',
	text  => 'включен'
    },
    1 => {
	color => '#DC0000',
	text  => 'выключен'
    }
);

my %COLOR_TRIGGER_PRIORITY = (
    0 => '#97AAB3',	#Not classified
    1 => '#7499FF',	#Information
    2 => '#FFC859',	#Warning
    3 => '#FFA059',	#Average
    4 => '#E97659',	#High
    5 => '#E45959'	#Disaster
);

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
    return 1;
}


#================================================================
sub send_to_zabbix
{
    my $json = shift;

    my $url = "http://" . ZABBIX_SERVER . "/api_jsonrpc.php";
    my $client = new JSON::RPC::Client;
    my $response = $client->call($url, $json);
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
	return 1;
    }
    print "Logout successful. Auth ID: $ZABBIX_AUTH_ID\n" if DEBUG;
    return 0;
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
    
    $data{'params'}{'sortfield'} = 'name'; #Possible values are: hostid, host, name, status.
    $data{'params'}{'sortorder'} = 'DESC'; #DESC or ASC
    
    $data{'auth'} = $ZABBIX_AUTH_ID;
    $data{'id'} = 1;
    
    my $response = send_to_zabbix(\%data);
    
    my $workbook = create_workbook('report');
    my $worksheet = create_worksheet($workbook, 'Hosts');
    $worksheet->freeze_panes(1, 0);
    $worksheet->autofilter('A1:H1');
	
    my $header_font = set_font($workbook, 2, 1, 'red', 12, 'Cambria', '#A0D8E5', 'center', 1);
	
    set_width_of_column($worksheet, 'A:A', 25);
    set_width_of_column($worksheet, 'B:B', 45);
    set_width_of_column($worksheet, 'C:C', 16);
    set_width_of_column($worksheet, 'D:D', 45);
    set_width_of_column($worksheet, 'E:E', 45);
    set_width_of_column($worksheet, 'F:F', 13);
    set_width_of_column($worksheet, 'G:G', 20);
    set_width_of_column($worksheet, 'H:H', 20);
    set_width_of_column($worksheet, 'I:I', 22);
    set_width_of_column($worksheet, 'J:J', 22);
    set_width_of_column($worksheet, 'K:K', 22);
    set_width_of_column($worksheet, 'L:L', 22);
    set_width_of_column($worksheet, 'M:M', 22);
    set_width_of_column($worksheet, 'N:N', 22);
    set_width_of_column($worksheet, 'O:O', 22);
    set_width_of_column($worksheet, 'P:P', 22);
    set_width_of_column($worksheet, 'Q:Q', 22);
    set_width_of_column($worksheet, 'R:R', 22);
    set_width_of_column($worksheet, 'S:S', 22);
    set_width_of_column($worksheet, 'T:T', 22);

    my @header_text = ('Сервер', 'Тест', 'Интервал, сек', 'Ключ теста', 'Описание теста',
		       'Статус', 'Период хранения истории (в днях)', 'Период хранения динамики изменений (в днях)',
		       'Триггер (чрезвычайный)', 'Описание триггера', 'Триггер (высокий)', 'Описание триггера',
		       'Триггер (средний)', 'Описание триггера', 'Триггер (предупреждение)', 'Описание триггера',
		       'Триггер (информация)','Описание триггера', 'Триггер (не классифицировано)', 'Описание триггера');

    foreach (0..$#header_text)
    {
	$worksheet->write(0, $_, $header_text[$_], $header_font);
    }
	
    parse_data($workbook, $worksheet, \@{$response->content->{'result'}});
	
    close_workbook($workbook);
}

#================================================================
sub parse_data
{
    my ($workbook, $worksheet, $ref_data) = @_;

    my $data_font = set_font($workbook, 1, 0, 'black', 12, 'Cambria', '#FFFFFF', 'left', 1);
    
    my $row = 1;
    foreach my $host(@{$ref_data})
    {
	foreach my $item(@{$host->{'items'}})
	{
	    write_to_worksheet($worksheet, $data_font, $host->{'name'}, $row, 0);
	    write_to_worksheet($worksheet, $data_font, $item->{'name'}, $row, 1);
	    write_to_worksheet($worksheet, $data_font, $item->{'delay'}, $row, 2);
	    write_to_worksheet($worksheet, $data_font, $item->{'key_'}, $row, 3);
	    write_to_worksheet($worksheet, $data_font, $item->{'description'}, $row, 4);
			
	    my $status_item_font = set_font($workbook, 1, 0, 'black', 12, 'Cambria', $TEST_STATUS{$item->{'status'}}{'color'}, 'left', 1);
		
	    write_to_worksheet($worksheet, $status_item_font, $TEST_STATUS{$item->{'status'}}{'text'}, $row, 5);
	    write_to_worksheet($worksheet, $data_font, $item->{'history'}, $row, 6);
	    write_to_worksheet($worksheet, $data_font, $item->{'trends'}, $row, 7);
			
	    my ($priority, $expression, $description) = get_trigger($item->{'itemid'});
			
	    if ($priority == 0) 
	    { 
		write_to_worksheet($worksheet, $data_font, $expression, $row, 8);
		write_to_worksheet($worksheet, $data_font, $description, $row, 9);
	    }
	    elsif ($priority == 1) 
	    { 
		write_to_worksheet($worksheet, $data_font, $expression, $row, 10);
		write_to_worksheet($worksheet, $data_font, $description, $row, 11);
	    }
	    elsif ($priority == 2) 
	    { 
		write_to_worksheet($worksheet, $data_font, $expression, $row, 12);
		write_to_worksheet($worksheet, $data_font, $description, $row, 13);
	    }
	    elsif ($priority == 3) 
	    { 
		write_to_worksheet($worksheet, $data_font, $expression, $row, 14);
		write_to_worksheet($worksheet, $data_font, $description, $row, 15);
	    }
	    elsif ($priority == 4) 
	    { 
		write_to_worksheet($worksheet, $data_font, $expression, $row, 16);
		write_to_worksheet($worksheet, $data_font, $description, $row, 17);
	    }
	    elsif ($priority == 5) 
	    { 
		write_to_worksheet($worksheet, $data_font, $expression, $row, 18);
		write_to_worksheet($worksheet, $data_font, $description, $row, 19);
	    }
	    $row++;
	}		
    }
}

#================================================================
sub get_trigger
{
    my $itemid = shift;
    my %data;

    $data{'jsonrpc'} = '2.0';
    $data{'method'} = 'trigger.get';
	
    my @triggers_params = ('triggerid', 'priority', 'expression', 'description', 'comments', 'status');
    $data{'params'}{'output'} = [@triggers_params];
    $data{'params'}{'itemids'} = $itemid;
	
    $data{'auth'} = $ZABBIX_AUTH_ID;
    $data{'id'} = 1;
	
    my $response = send_to_zabbix(\%data);
	
    my $priority;
    my $expression;
    my $description;
	
    foreach my $trigger(@{$response->content->{'result'}})
    {
	$priority = $trigger->{'priority'};
	$expression = $trigger->{'expression'};
	$description = $trigger->{'description'};
    }
	
    if (!defined $priority)
    {
	$priority = -1;
    }	
    return ($priority, $expression, $description);
}

#================================================================
#sub get_hosts_params
#{
#    my ($label, $ref_params, $ref_data) = @_;
#		
#    foreach my $data(@{$ref_data})
#    {					
#	foreach my $params(@{$ref_params})
#	{   
#	    print "$label: " . $data->{$params} . "\n";
#	}
#    }
#}

#================================================================
sub create_workbook
{
    my $file_name = shift;
	
    my $workbook  = Excel::Writer::XLSX->new("$file_name.xlsx");
    $workbook->set_properties(
	title    => 'Report about hosts',
	author   => 'Zabbix',
	comments => ''
    );
    return $workbook
}

#================================================================
sub create_worksheet
{
    my ($workbook, $worksheet_name) = @_;
	
    my $worksheet = $workbook->add_worksheet($worksheet_name);
    return $worksheet;
}

#================================================================
sub close_workbook
{
    my $workbook = shift;
    
    $workbook->close;
}

#================================================================
sub set_font
{
    my ($workbook, $border, $bold, $font_color, 
	$font_size, $font_type, $bg_color, $align, $wrap) = @_;
	
    my $font = $workbook->add_format(border => $border);
	
    $font->set_bold() if $bold;
    $font->set_color($font_color);
    $font->set_size($font_size);
    $font->set_font($font_type);
    $font->set_align($align);
    $font->set_align('vcenter');
    $font->set_bg_color($bg_color);
    $font->set_text_wrap() if $wrap;
    return $font;
}

#================================================================
sub write_to_worksheet
{
    my ($worksheet, $font, $text, $row, $col) = @_;
    
    $worksheet->write($row, $col, $text, $font);
}

#================================================================
sub set_width_of_column
{
    my ($worksheet, $column, $width) = @_;
    
    $worksheet->set_column($column, $width);
}

#================================================================
sub main
{
    if (zabbix_auth())
    {
	get_hosts();
	zabbix_logout();
    } 
    print "*** Done ***\n" if DEBUG;
}
