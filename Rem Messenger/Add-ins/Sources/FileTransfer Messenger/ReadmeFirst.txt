Secure File Transfer (SFT) by Tom Adelaar

All warranties and restrictions are given in the files coming with Secure File Transfer. It's just a kind of General Public License :)

Before compiling the program make sure you have the ebCrypt.dll reference right. This is done by going to Project -> References -> Browse (even when ebCrypt is already in your list!) -> Search for the path where you stored the ebCrypt.dll -> add it and make sure ebCrypt library is checked in you list and press OK. 

After this you can compile the Client and Server.

What is actually special about this Secure File Transfer? 
There are some nice features in it, which makes VB more suitable and stable for high speed file transfers. I included an extra TCP packet arrival buffer at the server side, so arriving packets aren't dropped when processing other packets in the buffer. The TCP packets in the buffer are processed in the spare time between arrival of successive TCP packets. 
At the Client side I added some TCP flow control check, which is often omitted in many VB file transfer examples I came across. Without this flow control, the server (or even the client) can get flooded with TCP packets. 
And of course, I added some crypto to the file transfer using the Incremental Rijndael (AES) Cipher from the ebCrypt.dll. 
You don't have to provide a key or pass-phrase (if you are lazy) to transfer files across a network, because I added a standard pass-phrase in the code. SFT can be used in combination with my Baboon Secure Chat application that I published previously on A1VBCode. 

At the moment this Secure File Transfer is suitable for internet (ADSL, Cable, etc..) For real high speed LAN throughput change in the Client code the TCP Blocksize to 8192 instead of 4096 bytes. 

If you have any questions, remarks or comments, please email.

Grtz,

Tom Adelaar
(TomAdelaar@hotmail.com)



