Ok this is the improved version of VBThreads. The previous version would not work correctly with VB5 due to the unattended execution not allowing Forms in the ThreadServer.EXE.....

Anyway, lets stop babbling and get on with it!

I have now written the project in VB5, If you open with VB6 then it will upgrade the project files to a VB6 schema.


_____________________________________________________________________________
TESTED ON:
_____________________________________________________________________________
120Mhz pentium Upto 350 Mhz Pentium with 16Meg ram on the 120 running Windows 95 and 128 Meg on the 350 running Win2000.No problems...(also tested on numerous NT boxes)



_____________________________________________________________________________
HOW TO GET IT WORKING:
_____________________________________________________________________________
1. open the VBIDE and load the ThreadServer project, Compile.(Do not run the server project from within the VBIDE. You will only get one thread running since the VBIDE will force this).

2. load the ThreadClient object.  Check that the references of the project include the previous ThreadServer that was compiled.

3. compile this and close the project!

4. Execute the ThreadClient app. Pressing the Start button will open a separate window with a counter. This counter will be in sync with the main window. If you press the Start thread again then another Thread will open and so on. The main window and the command buttons reflect on the last Thread to be started.

Keep on pressing and watch those threads.

If you seem to only get one thread to run at any one time then check the threading model of the ThreadServer project (Project properties/General Tab). This should be set to THREAD PER OBJECT.



_____________________________________________________________________________
EXPLANATION???
_____________________________________________________________________________
In ThreadServer there is an interface ITHREADABLE. Any object that you wish to be threadable implements that interface.(See the iterator class for an example). The iterator class calls the START method and passes itself has the ITHREADABLE object. This then sets up the object to start on a separate thread and then calls the ITHREADABLE.START method...The ball starts rolling from there!

The code is quite straight forward and it should not take too long for you to figure out.

If you want to create projects using separate threads then just use the ThreadServer project has a template (removing the iterator class example) and add your own ITHREADABLE class definitions.

Remember that if you add a FORM to the ThreadServer.EXE in VB5 then it immediatley sets unnatended execution to off and you will only get one thread running....

SO IN VB5 DONOT HAVE ANY OUTPUT COMING FROM THE THREADSERVER.EXE. 

(VB6 users are fine!)



_____________________________________________________________________________
IMPROVEMENTS:
_____________________________________________________________________________
1. The example that I have shown uses the RaiseEvent method to communicate back with the CLIENT. This can be slow depending on how frequent your thread communicates with the client. If your thread is to communicate extensivly, as with the iterator example, then you might want to implement a call back interface which will speed things up. If you want details on how to do this then donot hesitate to mail me at (c.saxton@dtn.ntl.com).

2. You can make the SERVER only allow a specific number of threads by choosing the THREAD POOL option instead of the THREAD PER OBJECT. Just set the option to the number of threads that you want to have open at any one time. Any threads that try to open once the MAX has been reached will wait for a thread to be freed before conitinuing.

3. The example that I have shown is very basic and it does not handle the Thread being closed if you close a thread window . (The window will close but the thread
keeps running to 4000!). You can easily make it do this by calling the stop thread method....I will leave that upto you...this was just intended as an example.....There is also no error handling routines....that I will also leave upto you...(It should not be too hard to put appropriate handlers where you want them).



_____________________________________________________________________________
CONCLUSION:
_____________________________________________________________________________
There are a number of advantages of using this method instead of the CreateThread API call. I find the API equivalent to be a pain in the back side when trying to use it in any useful, re-usable way. 

It is also very unstable. You cannot use the VBIDE to run a project when using the CreateThread function since it will crash the IDE. 

Using a separate ACTIVEX.EXE means that you can still debug the client without the IDE crashing since it runs OUT OF PROCESS. However....if you run the ThreadServer in an IDE and the ThreadClient in an IDE referencing the ThreadServer (bit of a mouth full) then all will work but only one thread will launch at a time, but it still won't crach though. IT is just more stable

The disadvantage is that you have got to have a separate EXE as well as your client EXE running but I have designed the ThreadServer to take up as little memory as I can so its not too bad. It runs fine on a 120Mhz Pentium with 16 Meg ram with about 20 threads just plowing along. I have also tested it on a 350Mhx 128 Meg machine and it runs like lightning.....I had about 50 threads and more just running fine....no probs.


Any improvements/problems then please mail me at (c.saxton@dtn.ntl.com)

If you are trying to do something in VB/VC/Java or whatever and you have a question then drop me a mail, I will help no matter what. I live, dream and think computers so the responce will be quick!