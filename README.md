# ClipSource Word Add-In
Add-in for MS Word which displays the source URL and citations of copied objects in combination with clipsource.py and clipsource_server.py (https://github.com/DavRet/ClipSource).

Features:
- Manage clipboard contents in MS Word, restore old clipboard data
- Displays the source URLs of copied objects
- Displays citations of copied object, if available
- Insert citations in bibliography
- Cite clipboard content and put citation in bibliography
- Insert copied image with caption which contains it's source
- Create footnotes containing the source of copied objects

How to run:
- Run clipsource.py (https://github.com/DavRet/ClipSource)
  - Extracts sources of copied objects and puts them on the clipboard
- Run clipsource_server.py (https://github.com/DavRet/ClipSource)
  - Enables communication with python script
- Open Visual Studio solution file (ClipSource.sln) and run the project there
  - After that, MS Word will open and you have to click the button at the right top to open the add-in
  - MS Word then asks if you want to permit access to the python server, click "yes" twice
