var documentation = require('documentation');
var fs = require('fs');


if (true) {

    fs.readFile('./files/source.asp', 'utf8', function (err,data) {
        if (err) {
            return console.log(err);
        }
        const newLine = "// ";
        var result = data

        result = result.replace(/\n/g, '\n' + newLine);
        result = result.replace(/<%/g, newLine);
        result = result.replace(/%>/g, newLine);
        result = result.replace(new RegExp(newLine + "'/\\*\\*", 'g'), '/**')
        result = result.replace(new RegExp(newLine + "'\\*/", 'g'), '*/')
        result = result.replace(new RegExp(newLine + "'[ |\t]*\\*", 'g'), ' *')

        fs.writeFile('./files/pretreatment.js', result, 'utf8', function (err) {
            if (err) return console.log(err);
        });
    });
}

if (false) {

    // https://github.com/documentationjs/documentation/issues/869
    // Version 9.1.1 and this happens again during the standard HTML generation. Adding the second option: {shallow: false} fixes it

    documentation.build(['./files/pretreatment.js'], {shallow: true})
        .then(documentation.formats.md)
        .then(output => {
            // output is a string of Markdown data
            fs.writeFileSync('./docs/test.md', output);
        })
        .then(()=> {
            fs.unlinkSync('./files/pretreatment.js')
        });

}

if (true) {
    var streamArray = require('stream-array');
    var vfs = require('vinyl-fs');

    documentation.build(['./files/pretreatment.js'], {shallow: true})
        .then(documentation.formats.html)
        .then(output => {
            streamArray(output).pipe(vfs.dest('./docs/test'));
        })
        .then(()=> {
            //fs.unlinkSync('./files/pretreatment.js')
        });
}