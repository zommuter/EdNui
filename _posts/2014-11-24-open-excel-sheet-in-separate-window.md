---
layout: post
title: "Open Excel sheet in separate window"
date: "2014-11-24"
---

While Excel 2013 finally [solved](http://blogs.office.com/2013/02/07/open-excel-workbooks-in-separate-windows-and-view-them-side-by-side/) this, previous versions of Excel provide no means to open sheets in separate windows - quite a nuisance if you want to compare two tables side-by-side. The typically posted workaround usually involves deactivating DDE, which does however cause more trouble than it solves. But as ever so often, a [StackExchange](http://stackexchange.com) site comes to the rescue:

[This post](http://superuser.com/a/277794/35237) gives a neat VBA-script to solve this problem in an almost comfortable way with only minimal one-time setup effort.

* First of all, start Excel and hit <kbd>Alt</kbd>+<kbd>F11</kbd> to open the VBA editor
* Right-click the `VBAProject (PERSONAL.XLSB)` and insert a module if none already exists. Give it a more sensible name than `Module1` while you're at it
* Double-click that module and paste the following code:  
  ```vbnet
    Sub OpenInNewInstance()
        ' http://superuser.com/a/277794/35237
        Dim objXLNewApp As Excel.Application
        Dim doc As String

        doc = ActiveWorkbook.FullName
        ActiveWorkbook.Close True

        Set objXLNewApp = CreateObject("Excel.Application")

        objXLNewApp.Workbooks.Open doc
        objXLNewApp.Visible = True
    End Sub
  ```
* Close the VBA-Editor after saving, but leave Excel open
* Right-click the quick access bar and choose to modify the same
* In the middle drop-down menu, select Macros and double-click `PERSONAL.XLSB!OpenInNewInstance`
* Click on the added entry in the rightmost list and modify the name to e.g. "Open in new Window" and pick an icon such as ![this one]({{ site.url }}/images/multiwin.png) for it.

Done. You now have a neat icon in your quick bar with which you can put the currently opened sheet into a separate window. The only down-side of this is that some of the DDE-functions for advanced copy-pasting fail.
