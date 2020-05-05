# MAdCaP.ps1

## Manage Analog Devices and Common Area Phones

Creating and administering Analog Devices and Common Area Phones in Lync and Skype for Business is quite tedious. (I have a terrible record for forgetting to assign a PIN to Common Area Phones&hellip;)

Inspired by a customer (and to the horror of a purist peer) I've created a PowerShell script that provides a GUI for the administration of both device types. I give you "MAdCaP.ps1" - short for "Manage Analog Devices & Common  Area Phones".

## The "New Object" tab

<img src="https://user-images.githubusercontent.com/11004787/81054933-0fc1e900-8f0b-11ea-8155-94a75400393f.png" alt="" width="600" />

## The "Existing Object" tab

<img src="https://user-images.githubusercontent.com/11004787/81054975-20725f00-8f0b-11ea-95cc-abd81dce4718.png" alt="" width="600" />

## Features

- Easily Create a new Analog Device or Common Area Phone, and optionally grant the relevant policies and PIN at the same time 

- Grant a Dial Plan, Voice Policy, Client Policy, Location Policy and/or User PIN to one or many already existing devices with a few clicks 

- Set a default OU to speed the creation of new objects and the editing of existing ones 

- Remembers your last-used Registrar Pool and SIP Domain - great for large (global) deployments with many pools and SIP domains 

- Filter your view of Existing Objects by OU, SIP Address, Line URI, Display Name or RegistrarPool 

- View the config of an existing Analog Device or Common Area Phone, and scrape it to the clipboard. (I use this a LOT - I copy the OU of an existing object prior to creating a new one) 

- Basic validity checks: the GO button won't be enabled until the mandatory parameters contain &lsquo;non-null' entries 

- Comprehensive input constraints: 

    - PIN and Line URI fields will only accept digits (the latter also permitting a "+") 

    - Display Name blocks the entry of restricted characters 

    - SIP URI can only be numeric, alpha, or limited punctuation 

    - OU and DN fields must contain at least "DN=", etc 

- All communication with Lync/SfB is phrased within error handlers to prevent the script crashing on bad user-input or system errors 

- Log to screen and optional text file 

- Works with Lync 2010, Lync 2013, Skype for Business 2015 & 2019

Please let me know if you think it's screaming out for any other features, or if something's broken when run against your deployment. I'm usually the quickest to reach via <a href="https://twitter.com/greiginsydney" target="_blank">Twitter</a>, otherwise by question here or on the <a href="https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones" target="_blank"> blog</a>.

## Limitations

- Your first visit to the Existing Object tab might be quite slow if you have a VAST deployment with lots of existing objects. Once you've set a default OU deeper into the forest (click the Filter button!) it will load much faster 

- If you're creating a New device there's usually a brief but noticeable lag between its creation and being able to then retrieve it and Grant policies. I've incorporated some handling that's meant to cater for this, but if you strike  any problems, just click to the Existing tab, Refresh, select your new device (which should have appeared by now), then re-apply the Policy settings 

- Related to the above, it seems it's not possible to set a PIN straight after creating a new Common Area Phone. (Others - <a href="http://social.technet.microsoft.com/Forums/en-US/ocsmanagement/thread/3b6f7edd-d46f-4f5b-b126-b351e5d54b57" target="_blank"> here</a> and <a href="http://blogs.perficient.com/microsoft/2011/04/how-to-create-lync-common-area-phones-in-bulk/" target="_blank"> here</a> - have tried and hit the same obstacle). The work-around is as per above 

- The "user-proofing" is comprehensive but not perfect. It'll capture errors to the on-screen log and optional text file if you're caught doing anything naughty (like specifying a name or number that's already in use, or a bad  OU) 

- If you name your new device "Error" it'll falsely trigger the error-handling & derail the script 

- It won't delete objects, or change the parameters of existing ones beyond the policies & PIN. I made a deliberate decision to omit the Remove functionality, and in so doing limiting the damage that someone could do if they were having a bad day 

### Revision History

#### v2.1 - 9th June 2018

- Rearranged calls to "handler_ValidateGo" to fix where the Go button wasn't lighting/going out 

- Corrected errors in the New Object DN and OU popup help text 

#### v2.0 - 29th April 2018

- Incorporated my version of Pat's "Get-UpdateInfo" so you'll be notified when I update it. (Credit: https://ucunleashed.com/3168) 

- Added test for AD module to prevent re-loading unnecessarily 

- Suppressed lots of "loading" noise from verbose output with "-verbose:`$false" 

- Stripped the "Tag:" name from the start of the relevant policies (kinda redundant, and was getting in the way of below) 

- Updated the Existing Objects tab: the Policies update in real-time to show the selected object's values 

- Replaced the "Browse" button on the Existing Objects tab with a new "Filter" button and form 

- Improved efficiency: 

    - The "Refresh" button on the Existing items tab (function "Update-DeviceList") now reads ADs & CAPs into separate global arrays 
    - "Update-Display" now just reads the item directly from the relevant array rather than re-querying 
    - "Update-DeviceList" no longer fires if the user Cancels or makes no change on the OU / Browse form 
    - "Grant-Policy" now checks the existing and new policy values & skips the commands that would make no change 

- Peppered "write-progress" throughout the loading process to help debugging 

- Added handling for "-debug" switch for in-depth debug display 

- Corrected tab order on the Existing items tab 

- Added "-ShowExisting" switch so you can launch with that tab selected 


There's more info and a full history over on <a href="https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones" target="_blank">my blog @ greiginsydney.com</a>&nbsp;- and you can find a list of <a href="https://greiginsydney.com/scripts/" target="_blank">all my other PowerShell utilities here</a>.


<br>

\- G.

<br>

This script was originally published at [https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones/](https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones/).



