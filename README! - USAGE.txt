WORKING PROCEDURE:
------------------
Well, it can monitor anything of course. What you have to do is to change your
previous proxy settings to new ones.

For example, to change the Internet Explorer's settings:
  * Go to Internet Options > Connections > LAN Settings
  * In the Proxy Servers section, specify your local computer name in Address
    or its better if you specify your own LAN IP previously provided to you by
    the Proxy administrator.
  * Change the Port to 8000 (this is for HTTP, for SOCKS connection set it to 1100).
    You can change these port number from the project's module. There is a constant
    which you can change to change the port numbers.

Finally, run the program and Turn On Proxy.

Now your Internet Explorer will connect to this program and this program
will further connect to your proxy (which is by default set to 192.168.0.1 with
HTTP port 8080 and SOCKS port 1080; you can also change these from the Module.)


SAMPLE USAGES:
--------------
The reason I created this program was because I lost my FTP server's password,
so I created this program and changed the proxy IP & port of my CuteFTP and then
my CuteFTP connected to my app. After then I found out what my password was; dumb
FTP does not encrypt passwords.

There are many MANY uses of this program, just think . . . . . :D



- Faraz Azhar (itz_faraz@hotmail.com)
Copyright: You're free to use this program and cannot spread it unless with my
permission of course. blah blah blah!