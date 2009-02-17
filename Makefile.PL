use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME              => 'Win32::WindowsUpdate',
    VERSION_FROM      => 'lib/Win32/WindowsUpdate.pm', # finds $VERSION
    PREREQ_PM         => {
      Win32::OLE => 0,
    }, # e.g., Module::Name => 1.1
    ($] >= 5.005 ?     ## Add these new keywords supported since 5.005
      (AUTHOR         => 'Dusty Wilson <dusty#megagram:com>') : ()),
);
