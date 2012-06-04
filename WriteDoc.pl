#!/usr/bin/perl -w
use strict;
use warnings;
use Win32::Word::Writer;
use XML::Parser;
use utf8;
use Encode qw{encode};
use feature "switch";
use FindBin;

my $bindir = $FindBin::Bin;
$bindir =~ s/\//\\\\/;

my $writters = DocWriters::Null->new(undef);

if (-e "doc_target/newdoc.doc") {
    unlink "doc_target/newdoc.doc" or die "can not delete file";
}

my $doc = Win32::Word::Writer->new();
$doc->Open("$bindir\\doc_helpers\\tamplatedoc.doc");
$doc->SaveAs("$bindir\\doc_target\\newdoc.doc");
$doc->MoveToEnd();
my @chapters = map { { basename => "chap$_" } } 1..12;
for (@chapters) {
    Chapter($_->{basename}.".xml");
}

$doc->SaveAs("$bindir\\doc_target\\newdoc.doc");
$doc = undef;

sub Chapter {
    my ($current) = @_;
    my $parser = XML::Parser->new(Handlers => { Start => \&H_Start, End => \&H_End, Char => \&H_Char });
    $parser->parsefile("source_xml/" . $current);
}

sub H_Start {
    my ($parser, $element_name, @attributs) = @_;
    my %attributs = @attributs;
    given ($element_name) {
        when ("b") {
            #$writters->Flush();
            #Selection.Font.Bold = wdToggle
            #Selection.Font.BoldBi = wdToggle
            #$doc->oSelection->Font->Bold($doc->rhConst->{wdToggle});
            #$doc->oSelection->Font->BoldBi($doc->rhConst->{wdToggle});
            #$writters->{PreserveInitalSpace} = 1;
            #$doc->SetBold(1);
            #$doc->ToggleBold();
            $writters->Write('*');
        }
        when ("english-line") {
            $writters->Flush();
            $doc->Write("\n");
            $writters = DocWriters::InsideText->new($writters);
            $doc->oSelection->LtrPara();
        }
        when ("a") {
            $writters->Flush();
            $writters = DocWriters::Link->new($writters);
            $writters->Create($attributs{href});
        }
        when (["html", "body"]) {
            # ignore
        }
        when ("title") {
            $writters = DocWriters::InsideText->new($writters);
            $doc->SetStyle(encode("cp1255", "כותרת 1"));
            $doc->oSelection->RtlPara();
        }
        when ("section") {
            $writters = DocWriters::InsideText->new($writters);
            $doc->SetStyle(encode("cp1255", "כותרת 2"));
            $doc->oSelection->RtlPara();
        }
        when ("text") {
            $writters = DocWriters::InsideText->new($writters);
            $doc->SetStyle(encode("cp1255", "בלוק עברית"));
            $doc->oSelection->RtlPara();
        }
        when ("code") {
            if ($writters->IsInsideText()) {
                $writters = DocWriters::CodeInsideText->new($writters);
            } else {
                $writters = DocWriters::CodeSection->new($writters);
            }
        }
        when ("br") {
            $writters->Flush();
            $doc->Write("\n");
        }
        when ("table") {
            $writters->Flush();
            $writters = DocWriters::Table->new($writters);
            $writters->Create($attributs{col},$attributs{row});
        }
        when (["td", "th"]) {
            if (exists $attributs{dir} and $attributs{dir} eq 'rtl') {
                $writters = DocWriters::InsideText->new($writters);
            } else {
                $writters = DocWriters::CodeInsideText->new($writters);
            }
        }

        default {
            print $element_name, "\n";
        }
    }
}

sub H_End {
    my ($parser, $element_name) = @_;
    given ($element_name) {
        when ("b") {
            $writters->Write('*');
            #$writters->Flush();
            #Selection.Font.Bold = wdToggle
            #Selection.Font.BoldBi = wdToggle
            #$doc->oSelection->Font->Bold($doc->rhConst->{wdToggle});
            #$doc->oSelection->Font->BoldBi($doc->rhConst->{wdToggle});
            #$doc->SetBold(0);
            #$doc->ToggleBold();
        }
        when ("english-line") {
            $writters->remove();
            $doc->Write("\n");
            $doc->oSelection->RtlPara();
        }
        when ("a") {
            $writters->remove();
        }
        when (["title", "section"]) {
            $writters->remove();
            $doc->Write("\n");
        }
        when ("br") {
            # did "\n" on start
        }
        when ("text") {
            $writters->remove();
            $doc->Write("\n");
        }
        when ("code") {
            $writters->remove();
        }
        when ("table") {
            #$doc->TableEnd();
            $writters->remove();
            $doc->oSelection->Rows->Delete();
            $doc->MoveToEnd();
        }
        when (["th", "td"]) {
            $writters->remove();
            $writters->Next();
        }
    }
}

sub H_Char {
    my ($parser, $str) = @_;
    $writters->Write($str);
}

package DocWriters;

sub new {
    my ($class, $prev) = @_;
    my $self =  bless { Prev => $prev }, $class;
    if (defined $prev) {
        $prev->Flush();
    }
    $self->OnStart();
    return $self;
}

sub OnStart {}
sub OnEnd {}
sub Flush {}

sub remove {
    my $self = shift;
    $self->Flush();
    $self->OnEnd();
    $writters = $self->{Prev};
}

sub IsInsideText {
    my ($self) = @_;
    my $p = $self;
    while (defined $p) {
        return 1 if ref($p) eq 'DocWriters::InsideText';
        $p = $p->{Prev};
    }
    return 0;
}

package DocWriters::Null;
our @ISA;
BEGIN { @ISA = ('DocWriters') }

sub Write {}

package DocWriters::AbstractTextWriter;
our @ISA;
BEGIN { @ISA = ('DocWriters') }
use Encode qw{encode};

sub ExtractEncodablePart {
    my ($encoding, $p_str) = @_;
    my $octats = encode($encoding, $$p_str, Encode::FB_QUIET);
    return $octats;
}

sub ExtractUnicodePart {
    my ($encoding, $p_str) = @_;
    my $unicode = '';
    while (length($$p_str) > 0) {
        my $chr = substr($$p_str, 0, 1);
        my $ochr = encode($encoding, $chr, Encode::FB_QUIET);
        if (length($ochr) > 0) {
            # the next char is an ascii
            last;
        } else {
            $unicode .= $chr;
            substr($$p_str, 0, 1, '');
        }
    }
    return $unicode;
}

sub try_encode {
    my ($self, $encoding, $string) = @_;
    local $@;
    my $octats;
    eval {
        $octats = encode($encoding, $string, Encode::FB_CROAK);
    };
    return $octats if defined $octats;
    print "Encoding error: |$string|\n\n";
    return encode($encoding, $string);
}

package DocWriters::InsideText;
our @ISA;
BEGIN { @ISA = ('DocWriters::AbstractTextWriter') }
use Encode qw{encode};

sub OnStart {
    my $self = shift;
    $self->{Buffer} = '';
    $self->{PreserveInitalSpace} = 0;
}

sub Write {
    my ($self, $str) = @_;
    $self->{Buffer} .= $str;
}

sub Flush {
    my $self = shift;
    my $str = $self->{Buffer};
    $self->{Buffer} = '';
    $str .= ' ';
    $str =~ s/[\s\r\n]+/ /g;
    if (not $self->{PreserveInitalSpace}) {
        $str =~ s/^\s+//;
    }
    $self->{PreserveInitalSpace} = 0;
    return if $str =~ m/^\s*$/;
    my $octats = $self->try_encode("cp1255", $str);
    $doc->Write($octats);
}

package DocWriters::CodeInsideText;
our @ISA;
BEGIN { @ISA = ('DocWriters::AbstractTextWriter') }
use Encode qw{encode};

sub Write {
    my ($self, $str) = @_;
    $doc->SetStyle(encode("cp1255", "קוד בתוך עברית"));
    $doc->oWord->Keyboard(1033); # moving to english
    $str =~ s/\n//g;
    $str =~ s/\s+/ /g;
    $str =~ s/^\s+//;
    $str =~ s/\s+$//;    
    return if $str =~ m/^\s*$/;
    $doc->Write("a");
    $doc->oSelection->TypeBackspace();
    $doc->Write($str);
    #$doc->oSelection->ClearFormatting();
    $doc->SetStyle(encode("cp1255", "עברית בתוך עברית"));
    $doc->oWord->Keyboard(1037); # moving to hebrew
    $doc->Write(encode("cp1255", "א"));
    $doc->oSelection->TypeBackspace();
    $self->{Prev}->{PreserveInitalSpace} = 1;
}

package DocWriters::CodeSection;
our @ISA;
BEGIN { @ISA = ('DocWriters::AbstractTextWriter') }
use Encode qw{encode};

sub OnStart {
    my $self = shift;
    $self->{Buffer} = '';
}

sub OnEnd {
    my $self = shift;
    my $str = $self->{Buffer};
    $str =~ s/^[\s\n]+//;
    $str =~ s/[\s\n]+$//;
    $doc->SetStyle(encode("cp1255", "קוד"));
    $doc->oSelection->LtrPara();
    $doc->oWord->Keyboard(1033);
    $doc->Write("a");
    $doc->oSelection->TypeBackspace();
    $doc->Write($self->try_encode("cp1255", $str)."a");
    $doc->oSelection->TypeBackspace();
    $doc->Write("\n");
    $doc->oSelection->RtlPara();
    $doc->oWord->Keyboard(1037);
}

sub Write {
    my ($self, $str) = @_;
    $self->{Buffer} .= $str;
}

package DocWriters::Table;
our @ISA;
BEGIN { @ISA = ('DocWriters') }

sub Create {
    my ($self, $x, $y) = @_;
	my $oTable = $doc->oDocument->Tables->Add(
			$doc->oSelection->Range,
			$y, $x, 
			$doc->rhConst->{wdWord9TableBehavior},
			$doc->rhConst->{wdAutoFitContent},
			);
    $self->{Table} = $oTable;
    $self->{wdCell} = $doc->rhConst->{wdCell}
}

sub Next {
    my $self = shift;
    $doc->oSelection->MoveRight({ Unit => $self->{wdCell} });
}

sub Write {
    # ignore
}

package DocWriters::Link;
our @ISA;
BEGIN { @ISA = ('DocWriters') }

sub Create {
    my ($self, $href) = @_;
    $self->{href} = $href;
    $self->{buffer} = '';
}

sub Write {
    my ($self, $str) = @_;
    $self->{Buffer} .= $str;
}

sub OnEnd {
    my $self = shift;

    #$sel->Application()->ActiveDocument()->Hyperlinks()->Add({ 'Anchor' => $sel->Range(), 
    # 	'Address' => "",
    #     'SubAddress' => $arg });
    $doc->oDocument()->Hyperlinks()->Add({
                                          'Anchor' => $doc->oSelection->Range,
                                          'Address' => $self->{href},
                                          'TextToDisplay' => $self->{Buffer} });
    $self->{Prev}->{PreserveInitalSpace} = 1;
}
