# AutoStructure
ESAPI Script for automatically creating helper structures.

The script adds a ring structure around the largest of the PTVs of the current plan or the
combined/merged PTVs if there is more than one PTV referenced in the plan name. The ring 
has a 3mm spacing between the largest PTV, 5mm between the first boost PTV and 7mm between the
second boost PTV.
The script creates cropped PTVs if the plan contains an integrated boost. The space between 
parent PTV and boost PTV is also 3mm. 
Helper structures for OARs are also created by cropping them 3mm from the largest or merged PTV.
The Script displays a warning message if any PRV overlaps with a PTV.

The script heavily relies on the naming conventions used in our department.
Example Breast:
PTV ids: PTV_1A, PTV_2A, PTV_3A
Boost id: PTV_1BA (=Site 1 Volume B in A)
Resulting plan id: c1A1BA2A3A Ma

Since PTVs are referenced in the plan name the script can create the helper structures needed 
by disassembling the name.
