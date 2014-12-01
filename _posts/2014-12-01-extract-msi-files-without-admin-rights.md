---
layout: post
title: "Extract MSI files without admin rights"
date: "2014-12-01"
---

Many great tools are unfortunately bundled in MSI files which require administrator rights for installation, despite the program running perfectly fine once extracted without requiring any actual installation. Fortunately, thanks to [this superuser.com answer](http://superuser.com/a/307679/35237) this simple one-liner can be used as the "Open with..." target for MSI files to extract its contents in most cases:

{% gist zommuter/c01ea8afc5b20adcb45e %}
