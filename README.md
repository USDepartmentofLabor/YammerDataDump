# YammerDataDump
A simple way to download your Yammer content to excel

Start by You scrolling all the way down to the bottom and "expand" all of the posts that are truncated.

Next, copy the whole Yammer from the webpage (with just CTRL+A and CTRL+C) and paste it into a excel file (CTRL+V).  

Then, just apply the macro code and viola... in just a couple seconds, you have an organized data dump from Yammer.

**Note:  All posts start with a line that has “FollowFollow” and the user name and comments always have “in reply to” and the user name. The difference in lines changes a little depending on if there’s comments, but they always have “Translated from English” before a post. I just built a little thing in VB.NET to do pull everything out from that logic (just because that’s what I use most), but it’s possible in any programming language (including Excel logic).  
