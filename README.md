# MAdCaP.ps1

## Manage Analog Devices and Common Area Phones

<p>Creating and administering Analog Devices and Common Area Phones in Lync and Skype for Business is quite tedious. (I have a terrible record for forgetting to assign a PIN to Common Area Phones&hellip;)</p>
<p>Inspired by a customer (and to the horror of a purist peer) I&rsquo;ve created a PowerShell script that provides a GUI for the administration of both device types. I give you &ldquo;MAdCaP.ps1&rdquo; &ndash; short for &ldquo;Manage Analog Devices &amp; Common  Area Phones&rdquo;.</p>
<h3>The "New Object" tab</h3>
<p><img id="198965" src="https://i1.gallery.technet.s-msft.com/madcapps1-a-gui-to-create-cbd2e3d8/image/file/198965/1/madcap-v2-new.png" alt="" width="624" height="544" /></p>
<p>&nbsp;</p>
<h3>The "Existing Object" tab</h3>
<p><img id="198966" src="https://i1.gallery.technet.s-msft.com/madcapps1-a-gui-to-create-cbd2e3d8/image/file/198966/1/madcap-v2-existing.png" alt="" width="624" height="544" /></p>
<p>&nbsp;</p>
<h3>Features</h3>
<ul>
<li>Easily Create a new Analog Device or Common Area Phone, and optionally grant the relevant policies and PIN at the same time </li>
<li>Grant a Dial Plan, Voice Policy, Client Policy, Location Policy and/or User PIN to one or many already existing devices with a few clicks </li>
<li>Set a default OU to speed the creation of new objects and the editing of existing ones </li>
<li>Remembers your last-used Registrar Pool and SIP Domain - great for large (global) deployments with many pools and SIP domains </li>
<li>Filter your view of Existing Objects by OU, SIP Address, Line URI, Display Name or RegistrarPool </li>
<li>View the config of an existing Analog Device or Common Area Phone, and scrape it to the clipboard. (I use this a LOT - I copy the OU of an existing object prior to creating a new one) </li>
<li>Basic validity checks: the GO button won&rsquo;t be enabled until the mandatory parameters contain &lsquo;non-null&rsquo; entries </li>
<li>Comprehensive input constraints: 
<ul>
<li>PIN and Line URI fields will only accept digits (the latter also permitting a &ldquo;+&rdquo;) </li>
<li>Display Name blocks the entry of restricted characters </li>
<li>SIP URI can only be numeric, alpha, or limited punctuation </li>
<li>OU and DN fields must contain at least "DN=", etc </li>
</ul>
</li>
<li>All communication with Lync/SfB is phrased within error handlers to prevent the script crashing on bad user-input or system errors </li>
<li>Log to screen and optional text file </li>
<li>Works with Lync 2010, Lync 2013 &amp; Skype for Business 2015 </li>
</ul>
<p>Please let me know if you think it&rsquo;s screaming out for any other features, or if something&rsquo;s broken when run against your deployment. I'm usually the quickest to reach via <a href="https://twitter.com/greiginsydney" target="_blank">Twitter</a>, otherwise by question here or on the <a href="https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones" target="_blank"> blog</a>.</p>
<h3>Limitations</h3>
<ul>
<li>Your first visit to the Existing Object tab might be quite slow if you have a VAST deployment with lots of existing objects. Once you've set a default OU deeper into the forest (click the Filter button!) it will load much faster </li>
<li>If you&rsquo;re creating a New device there&rsquo;s usually a brief but noticeable lag between its creation and being able to then retrieve it and Grant policies. I&rsquo;ve incorporated some handling that&rsquo;s meant to cater for this, but if you strike  any problems, just click to the Existing tab, Refresh, select your new device (which should have appeared by now), then re-apply the Policy settings </li>
<li>Related to the above, it seems it&rsquo;s not possible to set a PIN straight after creating a new Common Area Phone. (Others &ndash; <a href="http://social.technet.microsoft.com/Forums/en-US/ocsmanagement/thread/3b6f7edd-d46f-4f5b-b126-b351e5d54b57" target="_blank"> here</a> and <a href="http://blogs.perficient.com/microsoft/2011/04/how-to-create-lync-common-area-phones-in-bulk/" target="_blank"> here</a> &ndash; have tried and hit the same obstacle). The work-around is as per above </li>
<li>The &ldquo;user-proofing&rdquo; is comprehensive but not perfect. It&rsquo;ll capture errors to the on-screen log and optional text file if you&rsquo;re caught doing anything naughty (like specifying a name or number that&rsquo;s already in use, or a bad  OU) </li>
<li>If you name your new device &ldquo;Error&rdquo; it&rsquo;ll falsely trigger the error-handling &amp; derail the script </li>
<li>It won&rsquo;t delete objects, or change the parameters of existing ones beyond the policies &amp; PIN. I made a deliberate decision to omit the Remove functionality, and in so doing limiting the damage that someone could do if they were having a bad day </li>
</ul>
<h3>Revision History</h3>
<p>v2.1 - 9th June 2018</p>
<ul>
<li>Rearranged calls to "handler_ValidateGo" to fix where the Go button wasn't lighting/going out </li>
<li>Corrected errors in the New Object DN and OU popup help text </li>
</ul>
<p>v2.0 - 29th April 2018</p>
<ul>
<li>Incorporated my version of Pat's "Get-UpdateInfo" so you'll be notified when I update it. (Credit: https://ucunleashed.com/3168) </li>
<li>Added test for AD module to prevent re-loading unnecessarily </li>
<li>Suppressed lots of "loading" noise from verbose output with "-verbose:`$false" </li>
<li>Stripped the "Tag:" name from the start of the relevant policies (kinda redundant, and was getting in the way of below) </li>
<li>Updated the Existing Objects tab: the Policies update in real-time to show the selected object's values </li>
<li>Replaced the "Browse" button on the Existing Objects tab with a new "Filter" button and form </li>
<li>Improved efficiency: 
<ul>
<li>The "Refresh" button on the Existing items tab (function "Update-DeviceList") now reads ADs &amp; CAPs into separate global arrays </li>
<li>"Update-Display" now just reads the item directly from the relevant array rather than re-querying </li>
<li>"Update-DeviceList" no longer fires if the user Cancels or makes no change on the OU / Browse form </li>
<li>"Grant-Policy" now checks the existing and new policy values &amp; skips the commands that would make no change </li>
</ul>
</li>
<li>Peppered "write-progress" throughout the loading process to help debugging </li>
<li>Added handling for "-debug" switch for in-depth debug display </li>
<li>Corrected tab order on the Existing items tab </li>
<li>Added "-ShowExisting" switch so you can launch with that tab selected </li>
</ul>
<p>&nbsp;</p>
<p>There's more info and a full history over on&nbsp;<a href="https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones" target="_blank">my blog @ greiginsydney.com</a>&nbsp;- and you can find a list of <a href="https://greiginsydney.com/scripts/" target="_blank">all my other PowerShell utilities here</a>.</p>
<p>&nbsp;</p>
&nbsp;
<p>- Greig.</p>
