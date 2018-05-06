#!/usr/bin/perl -w

#perl report_hosts.pl --server 'https://host/api_jsonrpc.php' --user 'login' --pwd 'password' --file 'report.xlsx' --lang 'ru'

use strict;
use warnings;
use LWP::UserAgent;
use Getopt::Long qw(GetOptions);
use JSON qw(encode_json decode_json);
use Excel::Writer::XLSX;
use MIME::Lite;
use Data::Dumper;
use utf8;

use constant SMTP_SERVER => 'your_smtp_server';
use constant DEBUG       => 0; #0 - False, 1 - True

my %HOST_STATUS;
$HOST_STATUS{0}{'color'}        = '#00AA00';
$HOST_STATUS{0}{'text'}{'ru'}   = 'Включен';
$HOST_STATUS{0}{'text'}{'en'}   = 'Enabled';
$HOST_STATUS{1}{'color'}        = '#DC0000';
$HOST_STATUS{1}{'text'}{'ru'}   = 'Выключен';
$HOST_STATUS{1}{'text'}{'en'}   = 'Disabled';

my %TEST_STATUS;
$TEST_STATUS{0}{'color'}        = '#00AA00';
$TEST_STATUS{0}{'text'}{'ru'}   = 'Включен';
$TEST_STATUS{0}{'text'}{'en'}   = 'Enabled';
$TEST_STATUS{1}{'color'}        = '#DC0000';
$TEST_STATUS{1}{'text'}{'ru'}   = 'Выключен';
$TEST_STATUS{1}{'text'}{'en'}   = 'Disabled';

my %COLOR_TRIGGER_PRIORITY;
$COLOR_TRIGGER_PRIORITY{0} = '#97AAB3';     #Not classified
$COLOR_TRIGGER_PRIORITY{1} = '#7499FF';     #Information
$COLOR_TRIGGER_PRIORITY{2} = '#FFC859';     #Warning
$COLOR_TRIGGER_PRIORITY{3} = '#FFA059';     #Average
$COLOR_TRIGGER_PRIORITY{4} = '#E97659';     #High
$COLOR_TRIGGER_PRIORITY{5} = '#E45959';     #Disaster

my %HEADER_SETTINGS;
$HEADER_SETTINGS{'background'}  = '#A0D8E5';
$HEADER_SETTINGS{'font'}        = 'Cambria';
$HEADER_SETTINGS{'color'}       = 'red';
$HEADER_SETTINGS{'size'}        = 12;
$HEADER_SETTINGS{'align'}       = 'center';

my %DATA_SETTINGS;
$DATA_SETTINGS{'background'}  = '#FFFFFF';
$DATA_SETTINGS{'font'}        = 'Cambria';
$DATA_SETTINGS{'color'}       = 'black';
$DATA_SETTINGS{'size'}        = 12;
$DATA_SETTINGS{'aling'}       = 'left';

my %MESSAGE;
$TEST_STATUS{0}{'message'}{'ru'}   = 'Включен';
$TEST_STATUS{0}{'message'}{'en'}   = 'Включен';

my @HEADER_TEXT = (
    'Host name', 25,
    'Host status', 20,
    'Item', 45, 
    'Update interval', 11, 
    'Key', 45, 
    'Description', 45,
    'Status', 13, 
    'History storage period (in days)', 20, 
    'Trend storage period (in days)', 20,
    'Trigger (Disaster)', 22,
    'Description', 22,
    'Trigger (High)', 22, 
    'Description', 22,
    'Trigger (Average)', 22, 
    'Description', 22, 
    'Trigger (Warning)', 22, 
    'Description', 22,
    'Trigger (Information)', 22, 
    'Description', 22, 
    'Trigger (Not classified)', 22, 
    'Description', 22
);

my %HEADER_TEXT;
my $ZABBIX_SERVER;
my $ZABBIX_AUTH_ID;

binmode(STDOUT,':utf8');

main();

sub zabbix_auth
{
    my ($user, $pwd) = @_;

    my %data;

    $data{'jsonrpc'} = '2.0';
    $data{'method'} = 'user.login';
    $data{'params'}{'user'} = $user;
    $data{'params'}{'password'} = $pwd;
    $data{'id'} = 1;

    my $response = send_to_zabbix(\%data);
 
    $ZABBIX_AUTH_ID = get_result($response);
    do_debug('Auth ID: ' . $ZABBIX_AUTH_ID, 'SUCCESS');
}

sub zabbix_logout
{
    my %data;

    $data{'jsonrpc'} = '2.0';
    $data{'method'} = 'user.logout';
    $data{'params'} = [];
    $data{'auth'} = $ZABBIX_AUTH_ID;
    $data{'id'} = 1;

    my $response = send_to_zabbix(\%data);

    my $result = get_result($response);
    do_debug("Logout: $result", 'SUCCESS');
}

sub send_to_zabbix
{
    my $data_ref = shift;

    my $json = encode_json($data_ref);
    my $ua = create_ua();

    my $response = $ua->post($ZABBIX_SERVER,
                            'Content_Type'  => 'application/json',
                            'Content'       => $json,
                            'Accept'        => 'application/json'
    );

    if ($response->is_success)
    {
        my $content_decoded = decode_json($response->content);
        if (is_error($content_decoded))
        {
            do_debug('Error: ' . get_error($content_decoded), 'ERROR');
            exit(-1);
        }
        return $content_decoded;
    }
    else
    {
        do_debug('Error: ' . $response->status_line, 'ERROR');
        exit(-1);
    }
}

sub is_error
{
    my $content = shift;

    if ($content->{'error'})
    {
        return 1;
    }
    return 0;
}

sub get_result
{
    my $content = shift;

    return $content->{'result'};
}

sub get_error
{
    my $content = shift;

    return $content->{'error'}{'data'};
}

sub create_ua
{
    my $ua = LWP::UserAgent->new();

    $ua->ssl_opts(verify_hostname => 0, SSL_verify_mode => 0x00);
    return $ua;
}

sub colored
{
    my ($text, $color) = @_;

    my %colors = ('red'     => 31,
                  'green'   => 32,
                  'yellow'  => 33,
                  'blue'    => 34,
                  'magenta' => 35,
                  'cyan'    => 36,
                  'white'   => 37
    );
    my $c = $colors{$color};
    return "\033[" . "$colors{$color}m" . $text . "\e[0m";
}

sub do_debug
{
    my ($text, $level) = @_;

    if (DEBUG)
    {
        my %lev = ('ERROR'   => 'red',
                   'SUCCESS' => 'green',
                   'INFO'    => 'yellow'
        );
        print colored("$text\n", $lev{$level});
    }
}

sub parse_argv
{
    my $zbx_server;
    my $zbx_user;
    my $zbx_pwd;
    my $report_file;
    my $language;

    GetOptions('server=s'  =>  \$zbx_server,       #Zabbix server
               'user=s'    =>  \$zbx_user,         #User
               'pwd=s'     =>  \$zbx_pwd,          #Password
               'file=s'    =>  \$report_file,      #Name of file
               'lang=s'    =>  \$language          #Language

    ) or do { exit(-1); };

    if (defined($zbx_server))
    {
        return ($zbx_server, $zbx_user, $zbx_pwd, $report_file, $language);
    }
    else
    {
        do_debug('Option server requires an argument', 'ERROR');
        exit(-1);
    }
}

sub send_report
{
    my $report_file = shift;

    my $msg = MIME::Lite->new(
                                From    => 'zabbix-report@example.ru',
                                To      => 'test1@example.ru',
                                Subject => 'Report about hosts',
                                Type    => 'text/html',
                                Data    => '<h4>День Добрый!<br>Отчет во вложении</h4>'
    );
    $msg->attr('content-type.charset' => 'UTF-8');

    $msg->attach(
                 Path        => $report_file,
                 Type        => 'application/vnd.ms-excel',
                 Endcoding   => 'base64',
                 Disposition => 'attachment'
    );

    $msg->send('smtp', SMTP_SERVER, Debug => DEBUG);
}

sub delete_report
{
    my $report_file = shift;

    unlink($report_file);
}

sub get_hosts
{
    my ($workbook, $worksheet, $language) = @_;

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
    
    write_data($workbook, $worksheet, $language, \@{$response->{'result'}});
}

sub write_data
{
    my ($workbook, $worksheet, $language, $ref_data) = @_;

    my $data_font = set_font($workbook, 1, 0,
                            $DATA_SETTINGS{'color'}, 
                            $DATA_SETTINGS{'size'}, 
                            $DATA_SETTINGS{'font'},
                            $DATA_SETTINGS{'background'},
                            $DATA_SETTINGS{'align'},
                            1
    );
    
    my $row = 1;
    foreach my $host(@{$ref_data})
    {
	foreach my $item(@{$host->{'items'}})
	{
	    my $status_host_font = set_font($workbook, 1, 0, 'black', 12, 'Cambria', $HOST_STATUS{$host->{'status'}}{'color'}, 'left', 1);
	    my $status_item_font = set_font($workbook, 1, 0, 'black', 12, 'Cambria', $TEST_STATUS{$item->{'status'}}{'color'}, 'left', 1);

	    write_to_worksheet($worksheet, $data_font, $host->{'name'}, $row, 0);
	    write_to_worksheet($worksheet, $status_host_font, $HOST_STATUS{$host->{'status'}}{'text'}{$language}, $row, 1);
	    write_to_worksheet($worksheet, $data_font, $item->{'name'}, $row, 2);
	    write_to_worksheet($worksheet, $data_font, $item->{'delay'}, $row, 3);
	    write_to_worksheet($worksheet, $data_font, $item->{'key_'}, $row, 4);
	    write_to_worksheet($worksheet, $data_font, $item->{'description'}, $row, 5);
		
	    write_to_worksheet($worksheet, $status_item_font, $TEST_STATUS{$item->{'status'}}{'text'}{$language}, $row, 6);
	    write_to_worksheet($worksheet, $data_font, $item->{'history'}, $row, 7);

	   $row++;
	}		
    }
}

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

sub set_header
{
    my ($workbook, $worksheet, $language) = @_;

    my $header_font = set_font($workbook, 2, 1, 
                                $HEADER_SETTINGS{'color'}, 
                                $HEADER_SETTINGS{'size'}, 
                                $HEADER_SETTINGS{'font'}, 
                                $HEADER_SETTINGS{'background'}, 
                                $HEADER_SETTINGS{'align'}, 
                                1
    );

    my $col = 0;
    foreach (0..$#HEADER_TEXT)
    {
	if ($_ % 2 == 0)
        {
	    $worksheet->write(0, $col, $HEADER_TEXT[$_], $header_font);
	    $col++;
        }
        else
        {
            $worksheet->set_column($col -1, $col -1, $HEADER_TEXT[$_]);
        }
    }
    $worksheet->freeze_panes(1, 0);
    $worksheet->autofilter(0, 0, 0, (scalar @HEADER_TEXT -1) / 2);
}

sub write_trigger
{
    my ($worksheet, $data_font, $description, $priority, $expression, $row, $start_col) = @_;

    SWITCH: for ($priority)
    {
	/-1/ && do { last; };
	/5/ && do #Disaster
	{
	    write_to_worksheet($worksheet, $data_font, $expression, $row, $start_col);
	    write_to_worksheet($worksheet, $data_font, $description, $row, $start_col+1);
	    last;
	};
	/4/ && do #High
	{
	    write_to_worksheet($worksheet, $data_font, $expression, $row, $start_col+2);
	    write_to_worksheet($worksheet, $data_font, $description, $row, $start_col+3);
	    last;
	};	
	/3/ && do #Average
	{
	    write_to_worksheet($worksheet, $data_font, $expression, $row, $start_col+4);
	    write_to_worksheet($worksheet, $data_font, $description, $row, $start_col+5);
	    last;
	};
	/2/ && do #Warning
	{
	    write_to_worksheet($worksheet, $data_font, $expression, $row, $start_col+6);
	    write_to_worksheet($worksheet, $data_font, $description, $row, $start_col+7);
	    last;
	};	
	/1/ && do #Information
	{
	    write_to_worksheet($worksheet, $data_font, $expression, $row, $start_col+8);
	    write_to_worksheet($worksheet, $data_font, $description, $row, $start_col+9);	     
	    last;
	};
	/0/ && do #Not classified
	{
	    write_to_worksheet($worksheet, $data_font, $expression, $row, $start_col+10);
	    write_to_worksheet($worksheet, $data_font, $description, $row, $start_col+11);
	    last;
	};
    }
}

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
	
    foreach my $trigger(@{$response->{'result'}})
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

sub create_workbook
{
    my $file_name = shift;
	
    my $workbook  = Excel::Writer::XLSX->new($file_name);

    $workbook->set_properties(
	   title    => 'Report about hosts',
	   author   => 'Zabbix',
	   comments => ''
    );
    return $workbook
}

sub create_worksheet
{
    my ($workbook, $worksheet_name) = @_;
	
    my $worksheet = $workbook->add_worksheet($worksheet_name);
    return $worksheet;
}

sub close_workbook
{
    my $workbook = shift;
    
    $workbook->close;
}

sub write_to_worksheet
{
    my ($worksheet, $font, $text, $row, $col) = @_;
    
    $worksheet->write($row, $col, $text, $font);
}

sub set_center
{
    my $text = shift;

    my $width = `tput cols`;
    my $len = ($width / 2) - length($text) / 2;
    return ' ' x $len . "$text\n";
}

sub do_print
{
    my ($text, $level) = @_;

    my %lev = (
        'ERROR'   => 'red',
        'SUCCESS' => 'green',
    	'INFO'    => 'cyan'
    );
    print colored("$text\n", $lev{$level});
}

sub main
{
    my ($zbx_server, $zbx_user, $zbx_pwd, $report_file, $language) = parse_argv();

    $ZABBIX_SERVER = $zbx_server;

    zabbix_auth($zbx_user, $zbx_pwd);

    my $workbook = create_workbook($report_file);
    my $worksheet = create_worksheet($workbook, 'Hosts');

    set_header($workbook, $worksheet, $language);

    get_hosts($workbook, $worksheet, $language);

    close_workbook($workbook);

    send_report($report_file);

    delete_report($report_file);

    zabbix_logout();
}
