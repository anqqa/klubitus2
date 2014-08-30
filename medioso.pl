use LWP::UserAgent;
$filetopass="http://www.anqqa.com/klubitus/medioso.asp?haku=";

print "Content-Type: text/html\n";
print "Content-Length: 160\n\n";

$ua = new LWP::UserAgent;
$ua->agent("Mozilla/4.0");

$req = new HTTP::Request 'GET' => $filetopass;

$req->header('Accept' => 'text/html');
$res = $ua->request($req);

if ($res->is_success) {
     print $res->content;
} else {
     print "Error: " . $res->status_line . "\n";
}
