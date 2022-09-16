If you are not experienced with subclassing, TLBs and OCX development in general, it is advisable to use the compiled OCX instead.

You will experience an issue if you don't have the file OLEGuids.tlb registered (and in the same location).
It can be found under the folder control-source\CommonControls
Copy this file into Windows system directory (SysWOW64 in x64) and register it.