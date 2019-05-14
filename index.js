

/*
========================================================================================================================
=
========================================================================================================================
*/
if (false) {

    var documentation = require('documentation');
    var fs = require('fs');
    // https://github.com/documentationjs/documentation/issues/869
    // Version 9.1.1 and this happens again during the standard HTML generation. Adding the second option: {shallow: false} fixes it
    documentation.build(['./files/source.js'], {shallow: true})
        .then(documentation.formats.md)
        .then(output => {
            // output is a string of Markdown data
            fs.writeFileSync('./docs/output.md', output);
        });

}

/*
========================================================================================================================
= build
========================================================================================================================
Generate JavaScript documentation as a list of parsed JSDoc comments, given a root file as a path.
*/
if (false) {
    var documentation = require('documentation');

    documentation.build(['./files/source.js'], {
        // only output comments with an explicit @public tag
        // access: ['public']
    }).then(res => {
        // res is an array of parsed comments with inferred properties
        // and more: everything you need to build documentation or
        // any other kind of code data.
        console.log(res);

        // [ { description: { type: 'root', children: [Array], position: [Object] },
        //     tags: [ [Object], [Object], [Object], [Object] ],
        //     loc: SourceLocation { start: [Position], end: [Position] },
        //     context:
        // { loc: [SourceLocation],
        //     file: 'D:\\Web\\NodeJS\\NodeJS-Documentation\\files\\source.js',
        //     sortKey:
        //     '!D:\\Web\\NodeJS\\NodeJS-Documentation\\files\\source.js 00000009' },
        // augments: [],
        //     errors: [],
        //     examples: [],
        //     implements: [],
        //     params: [ [Object] ],
        //     properties: [],
        //     returns: [],
        //     sees: [],
        //     throws: [],
        //     todos: [],
        //     yields: [],
        //     kind: 'class',
        //     author: ': moi',
        //     name: 'Circle',
        //     members:
        // { global: [],
        //     inner: [],
        //     instance: [Array],
        //     events: [],
        //     static: [Array] },
        // path: [ [Object] ],
        //     namespace: 'Circle' } ]

    });
}

/*
========================================================================================================================
= formats
========================================================================================================================
Documentation's formats are modular methods that take comments and config as input and return Promises with results, like stringified JSON, markdown strings, or Vinyl objects for HTML output.

TypeError: Cannot read property 'parseExtension' of undefined
documentation.build(['./files/source.js'], {shallow: true})
*/
if (false) {
    var documentation = require('documentation');
    var streamArray = require('stream-array');
    var vfs = require('vinyl-fs');

    documentation.build(['./files/source.js'], {shallow: true})
        .then(documentation.formats.html)
        .then(output => {
            streamArray(output).pipe(vfs.dest('./docs'));
        });

}


/*
========================================================================================================================
= formats.markdown
========================================================================================================================
Formats documentation as Markdown.
*/

if (true) {

    var documentation = require('documentation');
    var fs = require('fs');

    documentation.build(['./files/source.js'], {shallow: true})
        .then(documentation.formats.md)
        .then(output => {
            // output is a string of Markdown data
            fs.writeFileSync('./docs/output.md', output);
        });

}

/*
========================================================================================================================
= formats.json
========================================================================================================================
Formats documentation as a JSON string.
*/

if (true) {
    var documentation = require('documentation');
    var fs = require('fs');

    documentation.build(['./files/source.js'], {shallow: true})
        .then(documentation.formats.json)
        .then(output => {
            // output is a string of JSON data
            fs.writeFileSync('./docs/output.json', output);
        });
}