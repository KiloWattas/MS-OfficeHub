# MS-OfficeHub
Repository that hosts links to official MS Office downloads and provides automated KMS activation tools.

This repository hosts text files containing links to download official MS office installers as well as python scripts that automates KMS activation process. Scripts require administrative privileges in order to execute corrently. If the script is not run as administrator, it will evoke UAC prompt and ask for elevation itself.

Scripts do not require any additional dependencies to be installed. It takes advantage of libraries that are already included in Python Standart Library: <code>os</code>,  <code>ctypes</code> and <code>sys</code>.

Scripts had been written in <code>Python 3.11.1</code> and tested on <code>Windows 10 22H2</code> machine.   

**KNOWN BUGS: Absolute path to system file must contain no spaces, otherwise the script will crash upon evoking UAC prompt. This will be fixed in the future.**

DISCLAIMER: The scripts for KMS activation proveded here are meant to be used as a template that should be modified for use with the license You purchased. The scripts intended purpose is to make extensive deployments to multiple machines easier. I do not take any responsibility for damages that rise from use or failure to use provided software.
