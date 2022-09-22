# Minimal XLSX

This project is to explore how minimal an .xlsx file can be without Excel complaining about it being broken. The "Content" folder contains all the files that will be zipped into an .xlsx file and opened. This allows modifying the files and quickly opening in Excel to verify the behavior.

Things to note:

The styles.xml file has quite a few elements that need to be defined. Most can exist with default values.

App.xml and Core.xml aren't needed. App carries version info which is useful. Subnote: the AppVersion must be a 3-part version (X.Y.Z), as a 4-part version will cause Excel to complain.

The worksheet has the opportunity for the most savings, since it tends to be the largest payload. A big reduction can be had by not explicitly defining the row/cell ref (`r` attribute), which is implied by the element order if not provided. By only writing the ref when there is a gap (null) in the cells, the size of this file can be reduced tremendously. 

Text data is normally written as a shared string value. Excel internally de-dupes the storage of string data, which would otherwise be very reptetive. In some scenarios, it might be better to not use the shared string table, if the value is known to be unique within the dataset. 
