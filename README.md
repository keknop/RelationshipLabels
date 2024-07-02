# What Does This Button Do?

This is an OCLC Connexion client macro that adds RDA relationship element IRIs to fields that contain one or more [PCC relationship labels](https://www.loc.gov/aba/rda/mgd/relationshipLabels/index.html). **Note: The LC-PCC relationship labels, like the rest of the LC-PCC Metadata Guidance for the official version of RDA, are not final and have not yet been approved for use in PCC cataloging.** Consider this a proof of concept for testing and feedback purposes.

Currently it only works on 1xx and 7xx name headings; title and name-title headings are not yet supported, though it will yell at you if you try to include $e/$j in the same field as $t. IRIs will be added in $4(s) at the end of the field.

**The macro will (should!):**

1. Match a label to a person, corporate body, or family relationship IRI based on the field tag and indicators. For example, the label "author" will be matched to the IRI for "author person" in 100 1_ or 0_, "author family" in 100 3_, or "author corporate body" in 110 or 111.
2. Flip original RDA relationship designators to PCC relationship labels when they are different (sponsoring body -> sponsor, etc.).
3. Flip labels to the correct form when the form varies based on the type of agent (issuing body -> issuing person if used in a 700 1_ field, for example).
4. Update spelling to match RDA/LC-PCC terminology for labels with multiple vernacular spellings (at present: honoree - > honouree; colorist -> colourist; draughtsman -> draftsman)
5. Prompt users for the WEMI domain to use for ambiguous labels (for example, "restorationist" covers both "restorationist [agent type] of expression" and "restorationist [agent type] of item").
6. Warn users if a relationship is not valid for a particular agent type (for example, "degree supervisor" is valid only for persons and "civil defendant" cannot be used for families).
7. Refrain from adding duplicate IRIs if that IRI is already present.

**The macro will not:**
 
1. Check to see if $4 IRIs already present in the field correspond to an existing label.
2. Re-control modified headings. There is no Connexion macro command to control a single heading, and calling "control all" every time is potentially extremely annoying.
3. Do anything at all for work/expession headings (yet).
4. Add or modify labels based on MARC relator codes (yet?).

## The disclaimers

The LC-PCC relationship labels, like the rest of the LC-PCC Metadata Guidance for the official version of RDA, are not final and have not yet been approved for use in PCC cataloging. Yes, I already said that. You can of course do whatever you like in your own local system, but refrain from using them in WorldCat or other shared databases until given the all-clear.

This macro is written for Connexion 3.0 or later. The macrobook itself will definitely not work in Connexion 2.63 or earlier; the macro itself may or may not, if modified, but that is left as an exercise for the determined reader.

This relies on a big chonky table of IRIs manually copy/pasted into a spreadsheet from the [RDA Registry](https://rdaregistry.info). There could be mistakes! In fact, given there are 202 labels corresponding to nearly 600 relationship IRIs in the table and the filter in the registry is not always as filtery as I might wish, I would be SHOCKED if at least one of them was not pointing to the wrong thing. Click the $4 links and make sure they're correct.

This is a fairly cumbersome pile of unoptimized code, written by someone who has not programmed anything in years, using an unfamiliar language that appears to be essentially "Visual Basic circa 1998, and you have to guess which VB features didn't exist yet because good luck finding any documentation on Softbridge Basic Language in 2024."  It involves a great deal of string manipulation and string comparison, which are always on the slow side of things codewise. It runs speedily on the relatively new and relatively-to-very beefy systems I have ready access to, but if you are running Connexion on a potato, I cannot guarantee it will be similarly performant.

# Installation

## Simple installation

You will need to download both the macrobook (RelationshipLabels.mbk) and the text file that contains the IRI mappings (RelationshipTable.txt) and save them both in your Connexion macro directory. By default, for Connexion 3.0 or later, that is C:\Users\[your user name]\AppData\Roaming\OCLC\Connex\Macros. In Windows File Exporer, typing %appdata% into the address bar and hitting enter will take you directly to the AppData\Roaming folder.

## Adding just one macro



## If your macros are not in the default location
If you have moved your macro directory to somewhere other than the default location (and would like to keep the the text file in the same place), you will need to edit the macro to look for the text file in the correct place. Go to *Tools > Macros > Manage...*, then expand the RelationshipLabels category, select RelationshipLabelAddIRI, and click *Edit...*. Search for "appdata" and locate the line:

```
sFileName = Environ$("APPDATA") & "\OCLC\Connex\Macros\RelationshipTable.txt"
```

You will need to replace the part beginning with "Environ$" with the path to wherever you have relocated your macros:

```
sFileName = "[full path to your Connexion macro directory]\RelationshipTable.txt"
```

You will then need to [click the "Check" button in the macro editor toolbar](https://help.oclc.org/Metadata_Services/Connexion/Connexion_client/Connexion_client_basics/Use_macros/Use_Connexion_client_macros/10Create_Connexion_client_macros#Check_macro_syntax). If all is well, you'll see a notification that the macro compiled successfully in the message bar at the bottom of the window. Click the save button.

# Use

Place your cursor in a 1xx or 7xx field with at least one LC-PCC relationship label or original RDA relationship designator and run the macro. General instructions on using macros in the Connexion client can be found [on the OCLC site](https://help.oclc.org/Metadata_Services/Connexion/Connexion_client/Connexion_client_basics/Use_macros/Use_Connexion_client_macros/20Work_with_Connexion_client_macros#Run_macros)

# Other Info

The labels, IRIs, and their associated WEMI domains are stored in a pipe-delimited list containing, in order, the label, the WEMI domain, and the IRIs for the corresponding person, corporate body, and family relationship elements. If a particular label/agent combination is not valid, in place of an IRI there is either a "USE:" reference, which punts the macro in the right direction, or an "ERR:" message that will be displayed to the user. The order of entries is technically arbitrary, since the macro will simply loop through the whole list until it finds a match (or runs out of things to check).

The macro relies on the text file having Windows-style line breaks to correctly count the number of entries. Given that Connexion is a Windows-only program that's probably not a concern, but something to be aware of I guess.

The spreadsheet file "relationship_mapping.xlsx" is not referenced by the macro at all, but is what was used to generate the entry strings for the text file.