package Win32::WindowsUpdate;

use strict;
use warnings;
use Win32::OLE qw(in);

our $VERSION = '0.01';

=head1 NAME

Win32::WindowsUpdate - Access to Windows Update functions

=head1 DESCRIPTION

Currently only provides the ability to see installed and available not installed Windows Updates.

The intention is to provide the ability to manage all features related to Windows Updates, including downloading,
installing, uninstalling (maybe), and configuring Windows Updates.

If you test this, please let me know your results.  It's not ready for production use, but any testing is
greatly appreciated.

=head1 EXAMPLE

  use Win32::WindowsUpdate;
  my $wu = Win32::WindowsUpdate->new;
  my $updates = $wu->updates;
  my $installed = $wu->installed;

  use Win32::WindowsUpdate;
  my $wu = Win32::WindowsUpdate->new({online => 0});
  my $updates = $wu->updates;

=head1 METHODS

=head2 new

  my $wu = Win32::WindowsUpdate->new;
  my $wu = Win32::WindowsUpdate->new({online => 0});

Creates the WindowsUpdate object.

=over

=item online

Defaults to C<1>.
Check online for updates.

=back

=cut

sub new
{
  my $class = shift;
  my $args = shift;
  my $self = {};

  $self->{online} = (defined($args->{online}) ? ($args->{online} ? 1 : 0) : 1); # online by default
  $self->{updateSession} = Win32::OLE->new('Microsoft.Update.Session') or die "ERROR creating Microsoft.Update.Session\n";
  $self->{updateSearcher} = $self->{updateSession}->CreateUpdateSearcher or die "ERROR creating CreateUpdateSearcher\n";
  $self->{updateSearcher}->SetProperty('Online', $self->{online});

  return bless($self, $class);
}

=head2 installed

  my $installed = $wu->installed;
  # .. or ..
  my @installed = $wu->installed;

Returns a list of installed updates.

=cut

sub installed
{
  my $self = shift;
  my $args = shift;

  return $self->updates({
    IsInstalled => 1,
    IsHidden => $args->{IsHidden},
  });
}

=head2 updates

  my $updates = $wu->updates;
  # .. or ..
  my @updates = $wu->updates;

  my $updatesHidden = $wu->updates({IsHidden => 1});

Returns a list of updates available for download and install.

=over

=item IsInstalled

Returns installed (if true) or not installed (if false) updates.
Defaults to C<0>.

=item IsHidden

Returns hidden (if true) or not hidden (if false) updates.
Defaults to C<0>.

=back

=cut

sub updates
{
  my $self = shift;
  my $args = shift;

  my $isInstalled = (defined($args->{IsInstalled}) ? ($args->{IsInstalled} ? 1 : 0) : 0);
  my $isHidden = (defined($args->{IsHidden}) ? ($args->{IsHidden} ? 1 : 0) : 0);
  my $query = $args->{query} || "IsInstalled = $isInstalled AND IsHidden = $isHidden";

  my $queryResult = $self->{updateSearcher}->Search($query);
  my $updates = $queryResult->Updates;

  my @updates;
  foreach my $update (in $updates)
  {
    my $info = {};

    $info->{UpdateId} = $update->Identity->UpdateID;
    $info->{Title} = $update->Title;
    $info->{Description} = $update->Description;
    $info->{RebootRequired} = $update->RebootRequired;
    $info->{ReleaseNotes} = $update->ReleaseNotes;
    $info->{EulaText} = $update->EulaText;
    $info->{EulaAccepted} = $update->EulaAccepted;

    push(@{$info->{Category}}, {Name => $_->Name, CategoryID=> $_->CategoryID}) foreach (in $update->Categories);
    push(@{$info->{MoreInfoUrl}}, $_) foreach (in $update->MoreInfoUrls);
    push(@{$info->{KBArticleID}}, $_) foreach (in $update->KBArticleIDs);
    push(@{$info->{SecurityBulletinID}}, $_) foreach (in $update->SecurityBulletinIDs);
    push(@{$info->{SupersededUpdateID}}, $_) foreach (in $update->SupersededUpdateIDs);

    push(@updates, $info);
  }

  return (wantarray ? @updates : \@updates);
}

=head1 TODO

=over

=item Provide ability to install specific updates.

=item Determine other necessary features. (email me with your requests)

=back

=head1 BUGS

B<REPORT BUGS!>
Report any bugs to the CPAN bug tracker.

=head1 COPYRIGHT/LICENSE

Copyright 2009 Megagram.  You can use any one of these licenses: Perl Artistic, GPL (version >= 2), BSD.

=head2 Perl Artistic License

Read it at L<http://dev.perl.org/licenses/artistic.html>.
This is the license we prefer.

=head2 GNU General Public License (GPL) Version 2

  This program is free software; you can redistribute it and/or
  modify it under the terms of the GNU General Public License
  as published by the Free Software Foundation; either version 2
  of the License, or (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see http://www.gnu.org/licenses/

See the full license at L<http://www.gnu.org/licenses/>.

=head2 GNU General Public License (GPL) Version 3

  This program is free software: you can redistribute it and/or modify
  it under the terms of the GNU General Public License as published by
  the Free Software Foundation, either version 3 of the License, or
  (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program.  If not, see http://www.gnu.org/licenses/

See the full license at L<http://www.gnu.org/licenses/>.

=head2 BSD License

  Copyright (c) 2009 Megagram.
  All rights reserved.

  Redistribution and use in source and binary forms, with or without modification, are permitted
  provided that the following conditions are met:

      * Redistributions of source code must retain the above copyright notice, this list of conditions
      and the following disclaimer.
      * Redistributions in binary form must reproduce the above copyright notice, this list of conditions
      and the following disclaimer in the documentation and/or other materials provided with the
      distribution.
      * Neither the name of Megagram nor the names of its contributors may be used to endorse
      or promote products derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
  WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
  PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
  ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
  LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
  OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
  IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

=cut

1;
