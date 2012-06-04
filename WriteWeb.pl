#!/usr/bin/perl -w
use strict;
use warnings;
use XML::Twig;
use Data::Dumper;
use utf8;

our ($prev_rec, $cur_rec, $next_rec);

main();

sub main {
    my @chapters = map { { basename => "chap$_" } } 1..12;
    BuildIndex(\@chapters);

    for (my $i=0; $i < @chapters; $i++) {
        my $prev = $i < 1 ? undef : $chapters[$i-1];
        my $next = $i >= $#chapters ? undef : $chapters[$i+1];
        local $prev_rec = $prev;
        local $next_rec = $next;
        local $cur_rec = $chapters[$i];
        Chapter($cur_rec);
    }
}

# each chapter record should have the following fields:
# filename, is_annex, title, sections=>[]

sub BuildIndex {
    my $chapters = shift;
    my $chapter_num = 1;
    my $i_writter = IndexWritter1->new();
    foreach my $rec (@$chapters) {
        $rec->{title} = undef;
        $rec->{sections} = [];
        $rec->{infile} = $rec->{basename} . ".xml";
        $rec->{outfile} = $rec->{basename} . ".html";
        $rec->{number} = $chapter_num;
        $rec->{"section_num"} = 1;
        
        my $twig=XML::Twig->new(twig_roots => 
                                   {
                                     title => sub {
                                        $rec->{title} = $_->text();
                                     },
                                     section => sub {
                                        push @{$rec->{sections}}, $_->text();
                                     },
                                    },
        );
        $twig->parsefile("source_xml/" . $rec->{infile}); # build it
        $rec->{extTitle} = "פרק" . " $chapter_num: " . $rec->{title};
        $i_writter->Chapter($rec->{extTitle}, $rec->{outfile});
        foreach my $sec (@{$rec->{sections}}) {
            $i_writter->Section($sec);
        }
        $chapter_num++;
    }
}

sub Chapter {
    my ($current) = @_;
    my $twig=XML::Twig->new(twig_handlers => 
                               { text => \&Handle_Text,
                                 "/html/body/code" => \&Handle_Code,
                                 "text/code" => \&Handle_code_inside_text,
                                 html => \&Handle_Main,
                                 "english-line" => \&English_line,
                                 title => \&Handle_Title,
                                 section => \&Handle_Section,
                                 description => \&Handle_Description,
                                 },
    );
    $twig->parsefile( "source_xml/" . $current->{infile} ); # build it
    $twig->print_to_file( "web_target/" . $current->{outfile} );  # output the twig
}

sub get_create_head {
    my ($twig) = @_;
    my $root = $twig->root();
    my $head = $root->first_child( 'head');
    if (not $head) {
        # head still not exists, create it
        $head = $root->insert_new_elt("first_child", "head");
    }
    return $head;
}

sub Handle_Description {
    my ($twig, $elm) = @_;
    my $head = get_create_head($twig);
    die "Description is too long in $cur_rec->{infile}" if length($elm->text()) > 164;
    $head->insert_new_elt("last_child", "meta", { description => $elm->text() });
    $elm->delete();
}

sub Handle_Title {
    my ($twig, $elm) = @_;
    my $title_text = $cur_rec->{extTitle};
    $elm->set_text($title_text);
    my $head = get_create_head($twig);
    $head->insert_new_elt("first_child", "title", {}, $title_text);
    $elm->set_tag("h1");
    BuildNavigation($elm, "after");
}

sub BuildNavigation {
    my ($elm, $position) = @_;
    my $last_in_hebrew = "הקודם :";
    my $next_in_hebrew = "הבא :";
    my $up_in_hebrew = "חזרה לתוכן העניינים";
    my $table = $elm->insert_new_elt($position, "table");
    my $row = $table->insert_new_elt("last_child", "tr");

    my $ref = sub {
        my ($text, $link) = @_;
        my $cell = $row->insert_new_elt("last_child", "td");
        $cell->insert_new_elt("last_child", "a", { href => $link }, $text);
    };

    if ($prev_rec) {
        $ref->("[$last_in_hebrew ".$prev_rec->{title}."]", $prev_rec->{outfile});
    }
    $ref->("[$up_in_hebrew]", "index.html");
    if ($next_rec) {
        $ref->("[$next_in_hebrew ".$next_rec->{title}."]", $next_rec->{outfile});
    }
}


sub Handle_Section {
    my ($twig, $elm) = @_;
    my $sec_num = join ".", $cur_rec->{number}, $cur_rec->{"section_num"};
    $elm->set_tag("h3");
    $elm->set_text( $sec_num . " " . $elm->text() ); # . '<a name="sec'.$cur_rec->{"section_num"}.'"></a>' );
    $elm->insert_new_elt("last_child", "a", { name => "sec".$cur_rec->{"section_num"} });    
    $cur_rec->{"section_num"}++;
}    

sub Handle_Text {
    my ($twig, $elm) = @_;
    $elm->set_tag("p");
    $elm->set_att(dir => "rtl");
    $elm->set_att(lang=>"he");
    my $div = $elm->wrap_in("div");
    $div->set_att(style => "padding-right: 30px");
}

sub Handle_Code {
    my ($twig, $elm) = @_;
    $elm->set_tag("pre");
    my $c = $elm->text();
    $c =~ s/^[\s\n]*//;
    $c =~ s/[\s\n]*$//;
    $elm->set_text($c);
    my $div = $elm->wrap_in("div");
    $div->set_att(dir => "ltr");
}

sub Handle_code_inside_text {
    my ($twig, $elm) = @_;
    $elm->set_tag("span");
    $elm->set_att(dir => "ltr");
    $elm->set_att(lang=>"en");
}

sub Handle_Main {
    my ($twig, $elm) = @_;
    $elm->set_att(dir => "rtl");
    BuildNavigation($elm, "last_child");

    #$elm->insert_new_elt("last_child", "br");
    my $copy_text1 = 'נכתב ע"י שמואל פומברג, כל הזכויות שמורות ©';
    my $copy_text2 = ' ראה פרק 1.5 לתנאי רשיון';
    my $span = $elm->insert_new_elt("last_child", "p", {}, $copy_text1.$copy_text2);
}

sub English_line {
    my ($twig, $elm) = @_;
    $elm->set_tag("div");
    $elm->set_att(dir=>"ltr");
    $elm->set_att(align=>"left");
}

package IndexWritter1;

# each chapter record contains:
#  { sections => [], file => '', name => '' }
# each section record contains:
#  { name => 'name' }

sub new {
    my ($class) = @_;
    my $self = { Chapters => [], current => undef };
    return bless $self, $class;
}

sub Chapter {
    my ($self, $chapter_name, $chapter_file) = @_;
    my $new_chap = { sections => [], file => $chapter_file, name => $chapter_name };
    push @{$self->{Chapters}}, $new_chap;
    $self->{current} = $new_chap;
}

sub Section {
    my ($self, $section_name) = @_;
    my $sec_rec = { name => $section_name };
    push @{$self->{current}->{sections}}, $sec_rec;
}

sub DESTROY {
    my $self = shift;
    my $chapter_num = 0;
    my $section_num = 0;
    my $ref = sub {
        my ($twig, $elm) = @_;
        $elm->set_tag("span");
        foreach my $chap_rec (@{$self->{Chapters}}) {
            $chapter_num++;
            $section_num = 0;
            my $h3 = $elm->insert_new_elt("last_child", "h3");
            $h3->insert_new_elt("last_child", "a", { href => $chap_rec->{file} }, $chap_rec->{name});
            foreach my $sec_rec (@{$chap_rec->{sections}}) {
                $section_num++;
                my $name = $section_num . " " . $sec_rec->{name};
                my $url = join '', $chap_rec->{file}, "#sec", $section_num;
                $elm->insert_new_elt("last_child", "a", { href => $url }, $name);
                $elm->insert_new_elt("last_child", "br");
            }
        }
    };

    my $twig = XML::Twig->new(twig_handlers => { insert => $ref } );
    $twig->parsefile( 'web_helpers/index_tamplate.xml' );
    $twig->print_to_file( "web_target/index.html" ); 
}
