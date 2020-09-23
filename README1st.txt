Updated: 11/10/99

Contributors to MDIChildMaker:
------------------------------
[YOUR NAME COULD BE HERE]


ORIGINAL CREDITS:
-----------------
 - Creator - Mark Einhorn (mje@cyberbiz.com) http://www.cyberbiz.com/ - PH: 773-338-3755
 - Based on the work of Galen Nickerson in the prjbinder.vbp project - http://www.nickerson.org/
 - Originally based on the work of Steve McMahon - http://www.vbaccelerator.com/


WHAT IS MDIChildMaker?
----------------------
THE PROBLEM:
------------
 VB allows forms to be instantiated from an ActiveX DLL...
 ...but does not allow MDI child forms to be instantiated from an ActiveX DLL!

THE SOLUTION: MDIChildMaker
---------------------------
 - MDIChildMaker handles the problem, but does not completely solve it, as there
   are known/bugs/issues (see below)
 - This took only about an hour to get working, so treat it as such!


OPEN SOURCE:
------------
 MDIChildMaker is an Open Source project. As such, we hope that the source code for MDIChildMaker will
 be readily available to all, and maintained by a community of VB programmers rather than individuals.

 To participate in the maintenance/enhancement of MDIChildMaker contact:
 EMAIL: mdicm@cyberbiz.com
 WEB: www.cyberbiz.com/mdicm/ - COMING SOON!


VERSION HISTORY:
----------------
 - 1.0 - 11/10/99 - Initial Release


FILE DESCRIPTIONS:
------------------
 - prjMDIChildMaker.vbp - contains MDI form sample, as well as all project forms, classes
   and modules to make CBMDIChildMaker work 
 - prjMDIForms.vbp - ActiveX DLL project with regular forms to be used an child forms
 - Group1.vbg - contains  prjMDIForms.vbp 


KNOWN BUGS/ISSUES WITH MDIChildMaker:
-------------------------------------
 - MDIChildMaker doesn't create real MDI child forms, but makes itself a child
   of the frmMDIChildContainer MDI child form - it would be nice to handle this 
   solution another way.
 - Form system messages don't get passed to the frmMDIChildContainer form - EX: ctrl-F6
 - doesn't automatically menus
 - Alt-Tab, causes the application to be first selected in the application list
   rather than the next application
 - When setting focus to an MDI form using MDIChildMaker, the first control
   on the form always get focus (rather than the previous control that had
   focus on the form)


THE FUTURE:
-----------
 - One day, hopefully, a future version of VB will solve this problem!
 - If you can fix any of the problems presented, by all means let me know...
   ...we'll fix the changes in the main distribution and post the changes in your name!


LICENSE:
--------
 - You have a royalty-free right to use, modify, reproduce and distribute all
   MDIChildMaker files.
 - By using any portion of MDIChildMaker you agree that CyberBiz, as well as all
   developers and distributors of MDIChildMaker have no warranty, obligations
   or liability for any MDIChildMaker related file (code, executable, etc.).


