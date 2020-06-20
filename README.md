# excel finance aggregate nodejs

1. install nodejs

https://nodejs.org/en/download/

2. clone or download this project

Click on `Clone or download` and select `Download ZIP`

3. Unzip the file to a new folder

Create a folder, and unzip the contents of the downloaded zip into the new folder.

4. In the new folder, create a folder called `input`.

Add the month folders into input.
The result directory should look like this:

```

main.js
package.json
.. (rest of the project files)
input/
    jan/
        a.xls
        b.xls
    feb/
        a.xls
        b.xls
```

5. Open command line in the directory.

6. Run this command

`npm install`

7. Run this command

`node main.js`

8. You should be able to see the result at output.xls
