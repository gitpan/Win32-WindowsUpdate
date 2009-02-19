package Win32::WindowsUpdate;

use strict;
use warnings;
use Win32::OLE qw(in);
use Inline::WSC;

our $VERSION = '0.02';

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

  die "Reboot first...\n" if $wu->rebootRequired;

  my @updates;
  foreach my $update ($wu->updates)
  {
    next unless ($update->{Title} =~ m/Security Update.*Windows XP/i); # filter for something
    push(@updates, $update); # push it into the updates list
  }

  sleep 1 while ($wu->installerBusy); # if Windows Installer is busy
  die "Reboot first...\n" if $wu->rebootRequired; # if whatever WI finished doing requires reboot
  $wu->install(@updates); # install your selected updates
  print "Installed updates want you to reboot...\n" if $wu->rebootRequired;

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
  $self->{updateSearcher}->{Online} = $self->{online};
  $self->{systemInfo} = Win32::OLE->new('Microsoft.Update.SystemInfo') or die "ERROR creating Microsoft.Update.SystemInfo\n";

  return bless($self, $class);
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

=head2 install

  my @updates;
  foreach my $update ($wu->updates)
  {
    printf("Available Update: %s %s\n", $update->{UpdateId}, $update->{Title});
    next unless ($update->{Title} =~ m/Security Update.*Windows XP/i);
    printf("Want to install: %s\n", $update->{Title});

    push(@updates, $update->{UpdateId});
    # .. or .. (both ways will work)
    push(@updates, $update);
  }

  $wu->install(@updates);

Install specified updates.
Provide an array of either C<update> (directly from C<< $wu->updates >>) or C<updateId>.
See example above for usage.

Currently implemented via VBScript (via L<Inline::WSC>).
Perl code exists in the module to do it without VBScript, but it doesn't work.
If you would like to help, please check out the source of this module to see what's going on there.

=cut

sub install
{
  my $self = shift;

  return undef unless scalar(@_); # no updates specified?  don't run.

  # using vbscript makes me sad... (see commented broken perl code after this sub)
  my $VBSCRIPT = <<VBSCRIPT;
    Sub vbInstall()
      Set vbSession = CreateObject("Microsoft.Update.Session")
      Set vbSearcher = vbSession.CreateUpdateSearcher()
      Set vbSysInfo = CreateObject("Microsoft.Update.SystemInfo")
      Set vbUpdateColl = CreateObject("Microsoft.Update.UpdateColl")
      Set vbSearchResult = vbSearcher.Search("IsInstalled = 0 AND IsHidden = 0")

      For i = 0 To vbSearchResult.Updates.Count - 1
        Set vbUpdate = vbSearchResult.Updates.Item(i)
VBSCRIPT

  # I'm sure there's a better way...
  foreach (@_)
  {
    my $uid = (ref($_) eq 'HASH' ? uc($_->{UpdateId}) : uc($_));
    warn("To install: $uid\n");
    $VBSCRIPT .= <<VBMATCHCODE;
        If UCase(vbUpdate.Identity.UpdateID) = "$uid" Then
          vbUpdate.AcceptEula
          vbUpdateColl.Add(vbUpdate)
        End If
VBMATCHCODE
  }

  $VBSCRIPT .= <<VBSCRIPT;
      Next

      Set vbDownloader = vbSession.CreateUpdateDownloader()
      vbDownloader.Updates = vbUpdateColl
      Set vbDownloadResult = vbDownloader.Download()

      Set vbInstaller = vbSession.CreateUpdateInstaller()
      vbInstaller.Updates = vbUpdateColl
      Set vbInstallationResult = vbInstaller.Install()
    End Sub
VBSCRIPT

  Inline::WSC->compile(VBScript => $VBSCRIPT);
  vbInstall();
}

# the perl way, but it doesn't seem to work... let me know how to fix it if you figure it out.
# if you fix it, include code that does accomplishes this: my $uid = (ref($_) eq 'HASH' ? uc($_->{UpdateId}) : uc($_));
#sub install
#{
#  my $self = shift;
#  my @updates = @_;
#
#  return undef unless scalar(@updates); # no updates specified?  don't run.
#
#  my $updatecoll = Win32::OLE->new('Microsoft.Update.UpdateColl') or die "ERROR creating Microsoft.Update.UpdateColl\n";
#  $updatecoll->Clear;
#
#  my $queryResult = $self->{updateSearcher}->Search("IsInstalled = 0 AND IsHidden = 0") or die "ERROR in query\n";
#  my $updates = $queryResult->Updates;
#
#  foreach my $update (in $updates)
#  {
#    my $updateID = $update->Identity->UpdateID;
#    next unless grep(m/^$updateID$/i, @updates);
#    $update->AcceptEula;
#    $updatecoll->Add($update);
#    printf("Adding %s %s to collection\n", $updateID, $update->Title);
#  }
#
#  printf("Updates to install: %d\n", $updatecoll->Count);
#
#  my $downloader = $self->{updateSession}->CreateUpdateDownloader;
#  $downloader->{Updates} = $updatecoll;
#  my $downloadResult = $downloader->Download;
#
#  my $installer = $self->{updateSession}->CreateUpdateInstaller;
#  $installer->{Updates} = $updatecoll; # <<<<<<<<<<<< problem is here... FIXME!
#  $installer->{AllowSourcePrompts} = 0;
#  $installer->{ForceQuiet} = 1;
#  my $installResult = $installer->Install;
#
#  use Data::Dumper;
#  print Dumper($updatecoll, $downloader, $downloadResult, $installer, $installResult);
#}

=head2 rebootRequired

  my $needToReboot = $wu->rebootRequired();

Returns bool.  True if reboot is required, false otherwise.

You should check this before and after you run C<install>.
If it's true before you install, you'll want to reboot before you install.
If it's true after you install, you'll want to reboot sometime.

=cut

sub rebootRequired
{
  my $self = shift;
  return ($self->{systemInfo}->RebootRequired ? 1 : 0);
}

=head2 installerBusy

  my $installerIsBusy = $wu->installerBusy();

Returns bool.  True if Windows Installer is busy, false otherwise.

If the installer is busy, you probably won't have any success at running C<install>.
You could just sit in a loop waiting for it to no longer be busy, but you'd want to check
C<rebootRequired> immediately before you run C<install> to be sure you won't want to reboot
first.

=cut

sub installerBusy
{
  my $self = shift;
  my $installer = $self->{updateSession}->CreateUpdateInstaller;
  return ($installer->IsBusy ? 1 : 0);
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
