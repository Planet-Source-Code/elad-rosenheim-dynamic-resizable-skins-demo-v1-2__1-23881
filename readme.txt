----------------------------------
"Dynamic" Resizable Skin Demo v1.1
----------------------------------
By Elad Rosenheim
mailto:eladro@barak-online.net

----------------------------------------------------------------------------
To see a fully-working application that uses this code, go to:
http://www.veoweb.com/users/kewlpad - home of kewlpAd, the skinnable notepad!
(free off course)
----------------------------------------------------------------------------

Hi there.
By now, we all know how to use bitmaps to create window
regions. But all these demos share one caveat:
The window is fixed - you can't resize it. That's ok for some 
apps, but sometimes you want to create a resizable window
in your app - The Playlist window in Winamp for example.

This demo shows you (one way) how to do that.
First of all, we can't rely on one bitmap. We need a number
of them:
One bitmap for each corner of the form (top-left, bottom-right...),
and a small bitmap for each of the form's sides - a bitmap that
can be seamlessly tiled to create a "side" as big as needed. 
In this demo I call these bitmaps "segments".

So here's how the window in constructed:

[1][2][2][2][3]
[4] Client  [5]
[4]  Area   [5]
[4]         [5]
[6][7][7][7][8]

[1] - Top left corner
[2] - Top horizontal segments - Can be horizantally tiled as many 
	times as needed.
[3] - Top right corner
[4] - Left vertical segments - Can be vertically tiled as many times 
	as needed.
[5] - Right vert. segments
[6] - Bottom-left corner
[7] - Bottom horiz. segments
[8] - Bottom-right corner

OK. So how do we create a region from all those bitmaps?
Obvisouly re-computing the whole region each time the window is 
resized is a really slow technique. And if we want to create three
regions based on bitmap [2], how do we make each region with a 
different X offset?

Fortunately, the API comes to the help (again). Turns out we can save a
region's definition to a binary buffer, then re-create it into a
new region. Moreover, we can offset an existing region by the
desired x,y offset, with one fast call.
Now we have the solution:
When the skin is loaded, we create regions for all the bitmaps and
store them.
Then, each time we need to create a new window region, We create 
copies of those regions, offsetting each region by the needed offset,
and combine 'em all.

We still have to tackle other problems, such as replacing the
standard window resizing mechanism with our own, which allow the window
to be resized only by fixed increments, which match the size of the
segments.
Sounds complicated? Start with looking at how the program runs, then
check out the LoadSkin() and MakeEdgesRegion() calls and the rest of 
the code.

Credits:
1. Steve McMahon of www.vbAccelerator.com for the clsBitmap code.
2. Myself... - I used clsDockHandler, a class previsouly published
   by me that handles docking to the edges of the desktop.
3. My friend Daniel Jacoby ("NevioTH") for the skins. If you wanna use 
   them - ask him for permission.

Off course, there's a lot more work to be done. For example - 
Most of the time you'll want to restrict the window from being resized
to less than a minimum size, depending in the controls you place
in the client area. And you'll want to resize some controls automatically
when the form size is changed. I leave it all to you.

Performace - 
The compiled version is a "smooth operator" (...) on my 366MHz
celeron. All possible optimization flags are already turned on.

The bigger the slices, the better the performance, because less slices are needed, but you need to find a good trade-off point between performace and the user's comfort.

------------------
Demo v1.1 Updates:

------------------

* Added support for skinnable control buttons (Minimize, Exit). 
Since the window is resizable, the buttons should also be moved when you resize the form, so there is support for positioning the buttons RELATIVE to the window sides.
Check out the clsButton class.

* Added a wrapper form to the project, so the app will have a normal taskbar button (Border-less forms cannot show an icon or have the 'minimize/restore' options in their taskbar button). The wrapper is invisible to the user, and only handles minimizing/restoring the main demo form.

 Note: If you want to see some great skins, go to the www.deviantart.com and browse the kewlpAd section.

Have a good time.

