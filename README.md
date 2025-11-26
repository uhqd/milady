# 012-Milady
This esapi script exports in a directory all the non-error documents of a patient in Aria. Open the patient and launch the script. Tada !

**Configuration** :

 - DocSettings.cs L24-40 : give a file with apikey credentials.
-  docFinder.cs L300 : export directory (must be visible from aria)
-  docFinder.cs L289-292 and 240: set type of documets. 
