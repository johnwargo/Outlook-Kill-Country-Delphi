Outlook-Kill-Country-Delphi
===========================

Deletes Outlook contact record country fields

For years, I've had this problem in Microsoft Outlook where Outlook will populate only the country field for select contacts. I've never done it purposefully, but somehow either because of some weird sync process or other reason, I had a bunch of contacts with no address, only a country set for them. Sigh
 
Anyway, this has been bugging me for some time and I've been too lazy to fix this manually. So, with all the Outlook integration code I've been doing lately, I decided I'd write an app that whacks the Country fields value (there are 4 of them: mailing, home, work and other) from every Outlook contact. This repository contains that code.

The code that accesses Outlook items came from [http://www.scalabium.com/faq/dct0121.htm](http://www.scalabium.com/faq/dct0121.htm). 

***

You can find information on many different topics on my [personal blog](http://www.johnwargo.com). Learn about all of my publications at [John Wargo Books](http://www.johnwargobooks.com).

If you find this code useful and feel like thanking me for providing it, please consider <a href="https://www.buymeacoffee.com/johnwargo" target="_blank">Buying Me a Coffee</a>, or making a purchase from [my Amazon Wish List](https://amzn.com/w/1WI6AAUKPT5P9).
