#!/usr/bin/perl -w
use strict;
use warnings;
use Win32::Word::Writer;
use XML::Parser;
use utf8;
use Encode qw{encode};
use feature "switch";

my $doc = Win32::Word::Writer->new();
$doc->Open("c:\\perl\\perlhebtut_xml\\words\\newdoc.doc");
#$doc->Open("c:\\perl\\perlhebtut_xml\\words\\tamplatedoc.doc");
#unlink "c:\\perl\\perlhebtut_xml\\words\\newdoc.doc" if -e "c:\\perl\\perlhebtut_xml\\words\\newdoc.doc";
#$doc->SaveAs("c:\\perl\\perlhebtut_xml\\words\\newdoc.doc");
$doc->MoveToEnd();

$doc->Write(encode("cp1255", "שלום שמואל "));
#$doc->oWord->Keyboard(1033);
$doc->oWord->Run("ToEnglish");
$doc->Write(encode("cp1255", "\$m "));
#$doc->oWord->Keyboard(1037);
$doc->oWord->Run("ToHebrew");
$doc->Write(encode("cp1255", "וזה הסוף "));

$doc->SaveAs("c:\\perl\\perlhebtut_xml\\words\\newdoc.doc");
$doc = undef;

