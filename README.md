Kapok.Report.Adomd
==================

A library adding Analytics Service client support for Kapok.Report by using the [ADOMD](https://learn.microsoft.com/en-us/analysis-services/client-libraries?view=asallproducts-allversions) library.

&nbsp;

Knowns issues
-------------
You might run into a GSSAPI error like shown [here](https://github.com/dotnet/runtime/issues/31579).
If that is the case, you need to run `sudo apt install -y --no-install-recommends gss-ntlmssp` to fix this.
