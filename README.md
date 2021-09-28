# vlookup-custom-formula-for-web-API-JSON
Along the way of our commercial B2B SaaS product development, that empowers frontline employees with intuitive text interface to poorly-accessible corporate data to inform and speed business decisions with significant benefits for enterprises, 

WE created a lot of useful and powerful software tools for testing apart of our main product. For example we've created VBA Excel function VLOOKUPWEB, that works same as VLOOKUP but for API JSON format responses.

https://user-images.githubusercontent.com/31845900/133978214-5707fc84-2b0a-4d94-8063-0f8280689f6e.mp4

=vlookupweb(1-link-to-api, 2-field-to-query, 3-time-out-in-seconds-between-calls, 4-token-if-needed, 5-body-for-post-or-empty-for-get-request) 

VLOOKUPWEB requires the following inputs: 
1. API link
2. Field name from JSON response you want to get data
3. Time-out parameters in seconds (0 for no time-out)
4. API token if Authorization required (if not leave empty)
5. Body parameter (if empty it would use GET method. If not empty Body will be sent as POST).

In order to add this function to your excel: 
1. Open excel 
2. Press Alt + f11 
3. create new module 
4. Copy-paste this code to new module 
5. Save and close VBA editor. 
6. You can use the function as typical formula similar to standard VLOOKUP

Enjoy! 

subscribe to our LinkedIn https://www.linkedin.com/company/nlsql-com
or Youtube https://www.youtube.com/channel/UC8KtzeNHxhLGVwiOCwvRBkg?sub_confirmation=1
in order to have even more great open-source tools, absolutely free!
