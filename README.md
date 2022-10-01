# Scientia Automation

This repository is for scripts which can help automate processes for Scientia Syllabus+. The scripts are written in VBScript, which means they can be scheduled using Windows Task Scheduler. 


## 1 To 1 Allocation
This script will automatically allocate students to "all group" teaching activities; these are the activities which all students enrolled on a module / course unit should attend.
In Syllabus+ terminology, these are the _Activity_ objects which do not share their _Activity Template_ with any other _Activity_ objects. Sometimes they are referred to as "1 to 1" activities, because the relationship between the _Activity Template_ and _Activity_ is 1:1.
This script will work as follows:
- It connects to a running Syllabus+ application whose registered Prog ID is "Splus".
- It creates an _Activity Template Group_ whose name includes the current date and time.
- It finds _Activity Template_ objects which meet the following three criteria and adds them to the group:
    - They have a _Module_ set.
    - There is only one _Activity_ attached.
    - The real size of that _Activity_ is smaller than the real size of the _Module_.
- It iterates through the members of the newly created _Activity Template Group_ and attaches all the _Student Sets_ of the _Module_ to the _Activity_. This has the effect of allocating them.
- Once completed, the script writes back the changes to the SDB.

Note that for this script to work, _Student Set_ objects corresponding to real students cannot have a size of zero.
