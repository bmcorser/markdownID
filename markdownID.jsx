//DESCRIPTION:Convert Markdown text or file to InDesign
// A Jongware script 04-Jun-2012

// Syntax based on the original Markdown:
//   http://daringfireball.net/projects/markdown/
// which bears the following copyright statement:
// > Copyright (c) 2004, John Gruber  
// > <http://daringfireball.net/>  
// > All rights reserved.
// ...
// although ID's data model needs some allowance here and there ...

if (parseFloat(app.version) < 6)
	getMarkdown();
else
	app.doScript (getMarkdown, ScriptLanguage.JAVASCRIPT,undefined,UndoModes.ENTIRE_SCRIPT, "Markdown Text");


function getMarkdown ()
{
	if (app.documents.length == 0)
	{
		alert ("Please create a document to insert Markdown text into first!");
	} else
	{
		tagset = findTagSet();
		if (app.selection.length > 0)
		{
			if (app.selection.length == 1 && app.selection[0].hasOwnProperty('baseline') && app.selection[0].length > 1)
			{
				datasource = new DataSource();
				datasource.setData (app.selection[0].contents);
				app.selection[0].remove();
				processMarkdown (datasource, tagset);
				return;
			}
		}

		folder = Folder.myDocuments;
		if (File.fs == "Windows")
			mdFile = folder.openDlg( 'Load Markdown document', "Markdown:*.md;*.mdown;*.markdown,Plain text:*.txt;*.text,All files:*.*", false);
		else
			mdFile = folder.openDlg( 'Load Markdown document', function(file) { return file instanceof Folder || (!(file.hidden) && file.name.match(/\.(txt|text|md|markdown|mdown|mdwn)$/i)); }, false );
		if (mdFile)
		{
			textsrc = getFromFile (mdFile);
			if (textsrc)
				processMarkdown(textsrc, tagset);
		}
	}
}

function findTagSet ()
{
	var i, f, l, tag, def, mapping = {
		'strong':'Markdown\\:STRONG',
		'em':'Markdown\\:EM',
		'strongem':'Markdown\\:STRONG+EM',
		'code':'Markdown\\:CODE',
		'hyperlink':'Markdown\\:HYPERLINK',

		'text':'Markdown\\:TEXT',
		'quote':'Markdown\\:BLOCKQUOTE',
		'pre':'Markdown\\:PRE',
		'h1':'Markdown\\:H1',
		'h2':'Markdown\\:H2',
		'h3':'Markdown\\:H3',
		'h4':'Markdown\\:H4',
		'h5':'Markdown\\:H5',
		'h6':'Markdown\\:H6',
		'bullet':'Markdown\\:BULLET',
		'numbered':'Markdown\\:NUMBERED',
		'indent':'Markdown\\:INDENT',
		'hrule':'Markdown\\:RULE',

		'th':'Markdown\\:TH',
		'td':'Markdown\\:TD',
	
	//	Table Style
		'table':'Markdown\\:Table',
	//	Cell Styles
		'header':'Markdown\\:Header',
		'body':'Markdown\\:Body'
	};

	f = app.activeDocument.textFrames.everyItem().getElements();
	for (i=0; i<f.length; i++)
	{
		if (f[i].contents.length > 0 && f[i].contents.substring(0,9).toLowerCase() == '#markdown')
		{
			l = f[i].parentStory.contents.split('\r');
			for (i=0; i<l.length; i++)
			{
				tag = l[i].match(/^(.+?)\s*=\s*(.+?)$/);
				if (tag && tag.length == 3)
				{
					tag[1] = tag[1].toLowerCase();
					if (mapping[tag[1]] != undefined)
						mapping[tag[1]] = tag[2].replace(/:/g, '\\:');
				}
			}
			break;
		}
	}
	return mapping;
}

function getFromFile (file)
{
	var datasource;
	var myFile = File(file);

	if (!myFile.open("r"))
	{
		alert ("Unable to open this file");
		return null;
	}
	myFile.encoding = 'binary';

	var skipBOM = 0;
	myFile.seek(0,0);

//	Can this be UTF-8?
	while (!myFile.eof)
	{
		block = myFile.readln();
		if ((block.charCodeAt(0) == 0 && block.charCodeAt(1) != 0) || (block.charCodeAt(0) == 0xfe && block.charCodeAt(1) == 0xff))
		{
			if (block.charCodeAt(0) == 0xfe)
				skipBOM = 2;
			myFile.encoding = 'utf16be';
			break;
		}
		if ((block.charCodeAt(0) != 0 && block.charCodeAt(1) == 0) || (block.charCodeAt(0) == 0xff && block.charCodeAt(1) == 0xfe))
		{
			if (block.charCodeAt(0) == 0xff)
				skipBOM = 2;
			myFile.encoding = 'utf16le';
			break;
		}
		m = block.match(/[\x80-\xFF]+/);
		if (m)
		{
			if (m[0].length == 1)
			{
			//	No way!
				myFile.encoding = 'latin1';
				break;
			}

			if (m[0].length >= 6)
			{
				myFile.encoding = 'latin1';
				if ((m[0].charCodeAt(0) & 0xfe) == 0xfc)
				{
					if ((m[0].charCodeAt(1) & 0xc0) == 0x80 && (m[0].charCodeAt(2) & 0xc0) == 0x80 && (m[0].charCodeAt(3) & 0xc0) == 0x80 && (m[0].charCodeAt(4) & 0xc0) == 0x80 && (m[0].charCodeAt(5) & 0xc0) == 0x80)
						myFile.encoding = 'utf8';
				}
				break;
			}
			if (m[0].length >= 5)
			{
				myFile.encoding = 'latin1';
				if ((m[0].charCodeAt(0) & 0xfc) == 0xf8)
				{
					if ((m[0].charCodeAt(1) & 0xc0) == 0x80 && (m[0].charCodeAt(2) & 0xc0) == 0x80 && (m[0].charCodeAt(3) & 0xc0) == 0x80 && (m[0].charCodeAt(4) & 0xc0) == 0x80)
						myFile.encoding = 'utf8';
				}
			}
			if (m[0].length >= 4)
			{
				myFile.encoding = 'latin1';
				if ((m[0].charCodeAt(0) & 0xf8) == 0xf0)
				{
					if ((m[0].charCodeAt(1) & 0xc0) == 0x80 && (m[0].charCodeAt(2) & 0xc0) == 0x80 && (m[0].charCodeAt(3) & 0xc0) == 0x80)
						myFile.encoding = 'utf8';
				}
			}
			if (m[0].length >= 3)
			{
				myFile.encoding = 'latin1';
				if ((m[0].charCodeAt(0) & 0xf0) == 0xe0)
				{
					if ((m[0].charCodeAt(1) & 0xc0) == 0x80 && (m[0].charCodeAt(2) & 0xc0) == 0x80)
					{
						myFile.encoding = 'utf8';
						if (m[0].charCodeAt(0) == 0xEF && m[0].charCodeAt(1) == 0xBB && m[0].charCodeAt(2) == 0xBF)
							skipBOM = 3;
					}
				}
				break;
			}
			if (m[0].length >= 2)
			{
				myFile.encoding = 'latin1';
				if ((m[0].charCodeAt(0) & 0xe0) == 0xc0)
				{
					if ((m[0].charCodeAt(1) & 0xc0) == 0x80)
						myFile.encoding = 'utf8';
				}
				break;
			}

		}
	}
//	Did we reach the end? No special characters then, but let's assume 'latin1', for completeness' sake...
	if (myFile.encoding == 'binary')
		myFile.encoding = 'latin1';

	myFile.seek(skipBOM,0);

	datasource = new DataSource();
	datasource.setData (myFile.read());
	return datasource;
}

function DataSource()
{
	this.text = '';
	this.length = 0;
	this.position = 0;
	this.eof = false;

	this.setData = function (txt)
		{
			this.text = txt.replace(/\r/g, '\n')+'\n';
			this.length = this.text.length;
			this.position = 0;
			this.eof = false;
		};

	this.readln = function ()
		{
			var prevpos = this.position;
			if (this.position >= this.length)
			{
				this.eof = true;
				return '';
			}
			var nextln = this.text.indexOf('\n', this.position);
			if (nextln == -1)
			{
				this.position = this.length;
				return this.text.substring (prevpos);
			}
			this.position = nextln+1;
			return this.text.substring (prevpos,nextln);
		}
	
};


function processMarkdown (source, tagset)
{
	var result, text, hyperlinkDefs;

	result = gatherParagraphs (source);
	text = result[0];
	hyperlinkDefs = result[1];

	text = textToStyle (text);
	text = resolveHyperlinks (text, hyperlinkDefs);
	createText (text);
}

function resolveHyperlinks (text, linkrefs)
{
	var i,l, hyperlinkDest = [];

	for (i=0; i<text.length; i++)
	{
		// Translate 'immediate' hyperlinks into regular syntax for the following part :(
		// .. this needs to check for 'proper' hyperlinks using GREP ... a black art ..
		//    .. so I'm taking some shortcuts here ..
		text[i] = text[i].replace(/(<0x005C><0x003C>((https?|ftp|mailto):\S+)<0x005C><0x003E>)/g, '[$2]($2)');

		//	Unhide a couple of common characters that may have got hidden for URLs:
		text[i] = text[i].replace(/<0x0022>/g, '"').replace(/<0x0027>/g, "'").replace(/<0x002d>/g, "-").replace(/<0x002e>/g, ".");

		if (text[i].indexOf('[') == -1)
			continue;
		text[i] = text[i].replace(/(\[([^\[\]]+?)\]) ?\[([^\[\]]*?)\]/g, function (full,text,innertext,linkref)
			{
				var UClink;
				if (linkref == '')
					UClink = innertext.toUpperCase();
				else
					UClink = linkref.toUpperCase();
				for (l=0; l<linkrefs.length; l++)
				{
					if (linkrefs[l][0] == UClink)
						return '['+innertext+']('+linkrefs[l][1]+' "'+linkrefs[l][2]+'")';
				}
				return full;
			}
		);

		// The 'destination' bit should be at the end? In reverse order? Oh okay then. (??)
		text[i] = text[i].replace (/\[([^\]]+)\]\s*\((\S+?)(\s+".+?")?\)/g, function(full, link, dest, title) {

			dest = dest.replace(/<cstyle:.*?>/g, '');
			dest = dest.replace(/^\\<?(.+)\\>?$/, '$1');
			var url = dest.replace(/([\/#:~])/g, '\\$1');
			url = url.replace(/<0x(....)>/g, function (match,code) { return String.fromCharCode(parseInt(code,16)); } );
			url = url.replace(/^\\<?(.+)\\>?$/, '$1');

		//	link = link.replace(/<0x(....)>/g, function (match,code) { return String.fromCharCode(parseInt(code,16)); } );
			link = link.replace(/<cstyle:.*?>/g, '');
			if (url.indexOf('@') >= 0)
			{
				if (!url.match(/^mailto\\:/))
				{
					url = 'mailto\\:'+url.replace(/@/g,'\\@');
				}
			}
			if (title == null)
				title = dest;
			else
				title = title.replace(/^\s*"(.+)"$/, '$1');
			
			title = title.replace(/<cstyle:.*?>/g, '');
			title = title.replace(/['"]/g, function (match) {
					match = "0000"+match.charCodeAt(0).toString(16);
					return "<0x"+match.substr(match.length-4)+">";
				} );

		//	We need the link length in *characters*
			linkLength = link.replace(/<0x....>/g, 'x').length;

		//	Hide a couple of character combos to prevent them being 'smartyfried'  ...
		//	url = url.replace(/-/g, '\x88').replace(/\./g, '\x89');
		//	title = title.replace(/-/g, '\x88').replace(/\./g, '\x89');

			title = title.replace(/<0x0020>/g, ' ');
		//	title = title.replace(/([^- A-Za-z0-9])/g, '\\$1');
		//	alert ("title: "+title+"\rlink: "+link+"\rurl: "+url); exit();
			hyperlinkDest.push ('<HplDestDfn:=<HplDestName:'+title+'><DestKey:'+String(hyperlinkDest.length)+'><HplDestUrl:'+url+'><Hid:0>>');
			return '<cstyle:'+tagset['hyperlink']+'><Hpl:=<HplName:'+title+'><HplDest:'+title+'><DestKey:'+String(hyperlinkDest.length)+'><CharStyleRef:'+tagset['hyperlink']+'><Hid:0><Brdrv:0><HplOff:0><HplLen:'+String(linkLength)+'>>'+link+'<cstyle:>';
			}
		);
	}
	text[text.length-1] += hyperlinkDest.join('');
	return text;
}

function textToStyle (source, linkrefs)
{
	var text = [], table = [], ln, inNumList = false;

	source.reverse();
	while (source.length > 0)
	{
		ln = source.pop();

		// Did we just finish building a table?
		if (table.length > 0 && ln[0] != 'TBH' && ln[0] != 'TBL')
		{
			text.push (createTable(table));
			table = [];
		}

		if (ln[0] == 'PRE')
		{
			// Need to 'unhide' UTF8 encoded characters
			ln[1] = ln[1].replace(/\xAE/g, '<').replace(/\xAF/g, '>');
		} else
		{
			if (ln[0] != 'LNK')
			{
			//	Process special characters
				ln[1] = accentEscape (ln[1], false);
	
			//	Hide stuff inside code fragments first
				// double ``encoded``
				ln[1] = ln[1].replace(/((``+)\s?)(.+?)(\s?\2)/g, function (_,start,__,text,end)
					{
						// .. need to do a dirty trick here to hide hyperlinks ..
						// .. these are recognized by their initial 'http|ftp|mailto' string ..
						text = text.replace(/<0x003C>([h|f|m])/g, function (_,match) {
								return '<0x003C>'+UCtoHiddenTag(match.charCodeAt(0));
							} );
						text = text.replace(/</g, '\xAE').replace(/>/g, '\xAF');
						text = text.replace(/[- |*_"'.\[\]`]/g, function (match)
							{
								return UCtoHiddenTag(match.charCodeAt(0));
							} );
						return '<cstyle:'+tagset['code']+'>'+text+'<cstyle:>';
					}
				);
				// single `encoded`
				ln[1] = ln[1].replace(/`(.+?)`/g, function (_,text)
					{
						// .. same as above ..
						text = text.replace(/<0x003C>([h|f|m])/g, function (_,match) {
								return '<0x003C>'+UCtoHiddenTag(match.charCodeAt(0));
							} );
						text = text.replace(/</g, '\xAE').replace(/>/g, '\xAF');
						text = text.replace(/[- |*_"'.\[\]]/g, function (match)
							{
								return UCtoHiddenTag(match.charCodeAt(0));
							} );
						return '<cstyle:'+tagset['code']+'>'+text+'<cstyle:>';
					}
				);
	
				// unhide special characters
			/*	Correction: uh no you don't. This might lead to non-ASCII characters in the text :(
				ln[1] = ln[1].replace(/<0x([0-9a-f]{4})>/g, function(_,code)
					{
						return String.fromCharCode(parseInt(code,16));
					}
				); */
	
				// unhide 'hidden tag' markers again
				ln[1] = ln[1].replace(/\xAE/g, '<').replace(/\xAF/g, '>');
			// ... now everything inside code fragments is 'protected'.

			// Hide characters inside hyperlinks from being smartyfried :(
				ln[1] = ln[1].replace (/\[([^\]]+)\]\s*\((\S+?)(\s+".+?")?\)/g, function(full, link, dest, title) {
					dest = dest.replace(/[\-."']/g, function (code)
						{
							return UCtoTag(code.charCodeAt(0));
						} );
					if (title == null)
						title = link;
					else
						title = title.replace(/\s+"(.+)"/, '$1');
					return '['+link+']('+dest+' <0x0022>'+title+'<0x0022>)';
					}
				);
	
			// Process default escape sequences
				ln[1] = ln[1].replace(/<0x005C><0x005C>([`*_{}\[\]()#+\-.!| '"])/g, function(_,code)
					{
						return UCtoTag(code.charCodeAt(0));
					} );
	
			//	Smart quotes
				// ... this is *ever* so slightly smarter than ID's own ...
				// (you're welcome to spot the difference!)
				ln[1] = ln[1].replace(/''/g, '"');
				ln[1] = ln[1].replace(/ ([_\*]*)"/g, ' $1<0x201C>');
				ln[1] = ln[1].replace(/"([_\*]*) /g, '<0x201D>$1 ');
				ln[1] = ln[1].replace(/(\w)"/g, '$1<0x201D>');
				ln[1] = ln[1].replace(/"(\w)/g, '<0x201C>$1');
				ln[1] = ln[1].replace(/(\S)"/g, '$1<0x201D>');
				ln[1] = ln[1].replace(/"/g, '<0x201C>');
		
				ln[1] = ln[1].replace(/ ([_\*]*)'/g, ' $1<0x2018>');
				ln[1] = ln[1].replace(/'([_\*]*) /g, '<0x2019>$1 ');
				ln[1] = ln[1].replace(/(\w)'/g, '$1<0x2019>');
				ln[1] = ln[1].replace(/'(\w)/g, '<0x2018>$1');
				ln[1] = ln[1].replace(/(\S)'/g, '$1<0x2019>');
				ln[1] = ln[1].replace(/'/g, '<0x2018>');
	
				// process default smart codes
				ln[1] = ln[1].replace(/---/g, '<0x2014>');
				ln[1] = ln[1].replace(/--/g, '<0x2013>');
				ln[1] = ln[1].replace(/\.\.\./g, '<0x2026>');
	
			//	Now we need to take care of table cells FIRST!
			//	(these may interfere with parsing of regular text attributes)
				if (ln[1].substring(0,1) == '|')
				{
					// First remove the very first and last |
					ln[1] = ln[1].replace(/^\|/, '').replace(/\|$/, '');
					ln[1] = ln[1].split('|');
				} else
					ln[1] = [ln[1]];
				numColumns = ln[1].length;
	
				for (i=0; i<numColumns; i++)
				{
					// Empty cell? Let one space remain:
					ln[1][i] = ln[1][i].replace(/^ +$/, '<0x0020>');
					// Save spaces left and right of | -- these determine center and right-align
					// Left: cull one space, then test the rest:
					ln[1][i] = ln[1][i].replace(/^ /, '');
					if (ln[1][i].match(/^ /))
					{
						ln[1][i] = ln[1][i].replace(/^ +/, '<0x0020>');
						// Right: same, but only if there was a left indent as well
						ln[1][i] = ln[1][i].replace(/ $/, '').replace(/ +$/, '<0x0020>');
					} else
						ln[1][i] = ln[1][i].replace(/ +$/, '');
	
	
					if (ln[1][i].indexOf('*') >= 0 || ln[1][i].indexOf('_') >= 0)
					{
						//  .. needs some special handling to allow "inside_a_word" notation
						ln[1][i] = ln[1][i].replace(/([0-9A-Z])_([0-9A-Z])/gi, '$1<0x005F>$2');

						// **strong** emphasized text
						ln[1][i] = ln[1][i].replace(/(\*\*|__)(((?!<cst).)+?)\1/g, '\xD1$2\xD0');
						// _regular_ emphasized text
						ln[1][i] = ln[1][i].replace(/(\*|_)(((?!<cst).)+?)\1/g, '\xD2$2\xD0');
			
						// Needs some effort for nested bold-in-italics or italics-in-bold :(
						ln[1][i] = ln[1][i].replace(/\xD1([^\xD0\xD2\xD3]*?)\xD2([^\xD0\xD2\xD3]*?)\xD0/g, '\xD1$1\xD3$2\xD1');
						ln[1][i] = ln[1][i].replace(/\xD2([^\xD0\xD2\xD3]*?)\xD1([^\xD0\xD2\xD3]*?)\xD0/g, '\xD2$1\xD3$2\xD2');
			
						ln[1][i] = ln[1][i].replace(/[\xD0\xD1\xD2\xD3]+([\xD0\xD1\xD2\xD3])/g, '$1');
			
						ln[1][i] = ln[1][i].replace(/\xD1/g, '<cstyle:'+tagset['strong']+'>');
						ln[1][i] = ln[1][i].replace(/\xD2/g, '<cstyle:'+tagset['em']+'>');
						ln[1][i] = ln[1][i].replace(/\xD3/g, '<cstyle:'+tagset['strongem']+'>');
						ln[1][i] = ln[1][i].replace(/\xD0/g, '<cstyle:>');
					}
				}
				if (numColumns > 1)
					ln[1] = ln[1].join('\t');
				else
					ln[1] = ln[1][0];
	
			//	Get rid of some more spaces
				ln[1] = ln[1].replace(/^ +/, '').replace(/ +$/, '').replace(/  +/g,' ');

			//	Unhide a couple of common characters that may have got hidden for URLs:
			//	ln[1] = ln[1].replace(/\x88/g, '-').replace(/\x89/g, '.');
			}
		}

		switch (ln[0])
		{
			case 'NEWLN': break;
			case 'HDR1': inNumList = false; text.push ('<pstyle:'+tagset['h1']+'>'+ln[1]); break;
			case 'HDR2': inNumList = false; text.push ('<pstyle:'+tagset['h2']+'>'+ln[1]); break;
			case 'HDR3': inNumList = false; text.push ('<pstyle:'+tagset['h3']+'>'+ln[1]); break;
			case 'HDR4': inNumList = false; text.push ('<pstyle:'+tagset['h4']+'>'+ln[1]); break;
			case 'HDR5': inNumList = false; text.push ('<pstyle:'+tagset['h5']+'>'+ln[1]); break;
			case 'HDR6': inNumList = false; text.push ('<pstyle:'+tagset['h6']+'>'+ln[1]); break;
			case 'TEXT': inNumList = false; text.push ('<pstyle:'+tagset['text']+'>'+ln[1]); break;
			case 'RULE': inNumList = false; text.push ('<pstyle:'+tagset['hrule']+'>'); break;
			case 'BULL': inNumList = false; text.push ('<pstyle:'+tagset['bullet']+'>'+ln[1]); break;
			case 'NUM':
				if (inNumList)
					text.push ('<pstyle:'+tagset['numbered']+'>'+ln[1]);
				else
					text.push ('<pstyle:'+tagset['numbered']+'><nmcfp:0>'+ln[1]);
				inNumList = true;
				break;
			case 'IND': text.push ('<pstyle:'+tagset['indent']+'>'+ln[1]); break;
			case 'PRE': text.push ('<pstyle:'+tagset['pre']+'>'+ln[1]); break;
			case 'BQ': text.push ('<pstyle:'+tagset['quote']+'>'+ln[1]); break;
			case 'LNK': text[text.length-1] += ln[1]; break;

			case 'TBH':
				table.push ("\x80"+ln[1]);
				break;
			case 'TBL':
				table.push (ln[1]);
				break;
			
			default:
				text.push ('<pstyle:'+tagset['text']+'>'+'<ct:Bold>('+ln[0]+')<ct:>'+ln[1]); break;
		}
	}
	return text;
}

function createTable (source)
{
	var text = '<pstyle:'+tagset['text']+'>', colWidths = [], numColumns = 0, numRows = source.length, numHdrs = 0, cells, row,cell, tmp, mrgd;
	
	for (row=0; row<source.length; row++)
	{
		// Remove Temporary marker for header rows
		if (source[row].substring(0,1) == '\x80')
		{
			numHdrs++;
			source[row] = source[row].substring(1);
		}

		cells = source[row].split('\t');
		if (cells.length > numColumns)
			numColumns = cells.length;

		cell = 0;
		while (cells.length > 0)
		{
			// Guesstimate column width, based on number of characters :(
			// This is not correct when there are merged cells!
			// .. well, at least it's better than the default 1" ..
			tmp = cells.shift();
			tmp = tmp.replace(/<cstyle:.*?>/g, '').replace(/<0x....>/g, 'x').length;
			if (colWidths[cell] == null || tmp > colWidths[cell])
				colWidths[cell] = tmp;
			cell++;
		}
	}

	// Create Table Tag
	text += '<tstyle:'+tagset['table']+'>';
	// Do not allow all rows set as header rows!
	//	.. InDesign does not like that *at all* ..
	if (numHdrs >= numRows)
		text += '<tStart:'+numRows+','+numColumns+':0:0>';
	else
		text += '<tStart:'+numRows+','+numColumns+':'+numHdrs+':0>';

	// Add Column Widths
	for (cell=0; cell<colWidths.length; cell++)
	{
		if (colWidths[cell] == null || colWidths[cell] < 2)
			colWidths[cell] = 1;
		text += '<coStart:<tcaw:'+(8+colWidths[cell]*7)+'>>';
	}

	// Add Rows
	for (row=0; row<numRows; row++)
	{
		text += '<rStart:>';

		// Surround each tab separated text with Cell Tags
		cells = source[row].split('\t');
		tmp = cells.length;
		while (cells.length > 0)
		{
			cell = cells.shift();
			// Merged cell? Then the next one(s) are empty.
			mrgd = 1;
			while (cells.length > 0 && cells[0].length == 0)
			{
				cells.shift();
				mrgd++;
			}
			// If it only contains a space, it was an empty cell
			if (cell == '<0x0020>') cell = '';

			// Set Cell Style Tag
			if (row < numHdrs)
				text += '<estyle:'+tagset['header']+'><clStart:1,'+mrgd+'><pstyle:'+tagset['th']+'>';
			else
				text += '<estyle:'+tagset['body']+'><clStart:1,'+mrgd+'><pstyle:'+tagset['td']+'>';

			// Test for a space left (= right aligned); if so, test right (= centered)
			if (cell.match(/^<0x0020>/))
			{
				if (cell.match(/<0x0020>$/))
					text += '<pta:Center>'+cell.match(/^<0x0020>(.+)<0x0020>$/)[1];
				else
					text += '<pta:Right>'+cell.match(/^<0x0020>(.+)$/)[1];
			} else
				text += cell;
			text += '<clEnd:>';

			// Fill with empty cells for merged ones
			while (mrgd-- > 1) text += '<clStart:><clEnd:>';
		}
	//	while (tmp++ < numColumns)
	//		text += '<estyle:'+tagset['header']+'><clStart:><clEnd:>'
		text += '<rEnd:>';
	}

	// End Table Tag
	text += '<tEnd:>';

	return text;
}

function gatherParagraphs (source)
{
	var lines = [], text = [], hyperlinkDefs = [], block, type, ln, match, level;

	while (!source.eof)
	{
		ln = source.readln();

		// Replace Tabs with 4 spaces
		ln = ln.replace(/\t/g, '    ');

		// Ensure blank lines are actually empty
		ln = ln.replace (/^ +$/,'');

		// Make sure lines end with 2 spaces OR none (as only 2 spaces is 'special')
		if (ln.match(/  $/))
			ln = ln.replace(/ +$/, '  ');
		else
			ln = ln.replace(/ $/, '');

		// Skip blank lines at the start and multiple blank lines
		if (ln == '')
		{
			if (lines.length == 0 || lines[lines.length-1] == '')
				continue;
			lines.push ('');
			continue;
		}

		// Remove properly formatted link references right away.
		// Totally disregard them from input-- do not even count as 'blank line'.
		// It *does* mean these link defs get discarded even if they are not
		// used. Shouldn't be a problem, though. You ought to know what you are
		// doing, right?
		// 		[foo]: http://example.com/  "Optional Title Here"
		//		[foo]: http://example.com/  'Optional Title Here'
		//		[foo]: http://example.com/  (Optional Title Here)
		//	Do not attempt to check all of the above in one expression! For some
		//	reason it's **hideously** slow!
		if (match = ln.match(/^ {0,3}\[([^\]]+)\]:\s+(\S+)\s*$/) ||
			match = ln.match(/^ {0,3}\[([^\]]+)\]:\s+(\S+)\s*(".+?")?\s*$/) ||
			match = ln.match(/^ {0,3}\[([^\]]+)\]:\s+(\S+)\s*('.+?')?\s*$/) ||
			match = ln.match(/^ {0,3}\[([^\]]+)\]:\s+(\S+)\s*(\(.+?\))?\s*$/))
		{
		//	Save the reference, the URL, and its title if there is one, or the original text otherwise.
			if (match[3] == null)
				hyperlinkDefs.push ([ match[1].toUpperCase(), match[2], match[1] ]);
			else
				hyperlinkDefs.push ([ match[1].toUpperCase(), match[2], match[3].substr(1,match[3].length-2) ]);
			continue;
		}

		// Early Detection needed for = and - underlined headers!
		if ((match = ln.match(/^[-=]+ *$/)) && lines.length > 0 && lines[lines.length-1] != '')
		{
			if (match[0][0] == '=')
				lines[lines.length-1] = '# '+lines[lines.length-1].replace(/^ +/, '');
			else
				lines[lines.length-1] = '## '+lines[lines.length-1].replace(/^ +/, '');
			continue;
		}

	//	Hide Tagged Text special characters < > /
		ln = ln.replace(/</g, '\x80' );
		ln = ln.replace(/>/g, '<0x005C><0x003E>' );
		ln = ln.replace(/\x80/g, '<0x005C><0x003C>' );
		ln = ln.replace(/\\/g, '<0x005C><0x005C>' );

	//	Resolve UTF8 characters into ASCII encoding
		ln = ln.replace(/[\u0080-\uFFFF]/g, function(match) { return UCtoHiddenTag(match.charCodeAt(0)); } );

	//	'Hide' code sequences that might mess up proper recognition. At this
	//	point we do NOT want to interpret them, just hide!
	//	Hiding is done by replacing them with their hex representation '<0x00XX>'
		/* List is from http://daringfireball.net/projects/markdown/syntax
		   "\   backslash
			`   backtick
			*   asterisk
			_   underscore
			{}  curly braces
			[]  square brackets
			()  parentheses
			#   hash mark
			+   plus sign
			-   minus sign (hyphen)
			.   dot
			!   exclamation mark" */
		// [JW]: Added | because this indicates a table */
		//	     Added space because you might want to use more than one
		//		 Added ' and " to get literals instead of 'smart' ones
		//		 The backslash itself is not necessary here because it was already culled above
		ln = ln.replace(/<0x005C><0x005C>([`*_{}\[\]()#+\-.!| '"])/g, function(match, code) { return UCtoTag(code.charCodeAt(0)); } );

	//	Jongware addition: special character escape codes
		// Also just hide them, do not process yet.
	//	ln = accentEscape (ln, true);

		lines.push (ln);
	}
	// Make sure we end with at least ONE blank line
	lines.push('');

	block = '';
	type = 'TEXT';

	lines.reverse();

	while (lines.length > 0)
	{
		ln = lines.pop();
		
		// Blank line?
		if (ln == '')
		{
			if (block != '')
				text.push ([type, block]);

			block = '';

			// Not necessary at start or after a header, rule, or bullet/number list item:
			if (text.length == 0 || text[text.length-1][0] == 'RULE' || text[text.length-1][0].substr(0,3) == 'HDR' || text[text.length-1][0] == 'BULL' || text[text.length-1][0] == 'NUM')
				continue;
			text.push (['NEWLN', 'xxx']);
			continue;
		}

		// The following line starters will all force-end
		// a previous paragraph:

		// Rule (needs to be tested first, because '*'
		// also may start a bulleted list!)
		if (ln.match(/^ {0,3}([-*_])( *\1){2,} *$/))
		{
			if (block != '')
				text.push ([type, block]);
			block = '';
			text.push (["RULE", '---']);
			continue;
		}

		// Numbered list: 1.(space)
		if (ln.match(/^ {0,3}\d+\. +\S/))
		{
			if (block != '')
				text.push ([type, block]);
			block = ln.replace(/^ *\d+\. */, '');
			type = "NUM";
			level = 0;
			continue;
		}

		// Bulleted list: * or + or -(space)
		if (ln.match(/^ {0,3}[-*+] +\S/))
		{
			if (block != '')
				text.push ([type, block]);
			block = ln.replace (/ *[-*+]\s+/, '');
			type = "BULL";
			level = 0;
			continue;
		}

		// Header: run of either # or =
		if (match = ln.match(/^(#{1,6})\s*\S/) || match = ln.match(/^(={1,6})\s*\S/))
		{
			if (block != '')
				text.push ([type, block]);
			block = '';
			// Discard start and end # or =
			text.push ([ "HDR"+match[1].length, ln.replace(/^([#=])\1{0,5}\s*(.+?)\s*\1* *$/, '$2') ] );
			continue;
		}

		// Table row? Must start with | for that
		if (ln.match(/^ {0,3}\|/))
		{
			if (block != '')
			{
				text.push([type,block]);
				block = '';
			}
			ln = ln.replace(/^ +/,'').replace(/\| +$/,'|');
			ln = ln.replace(/ +$/, '');
			if (!ln.match(/\|$/)) ln = ln + '|';
			if (ln.match(/^(\|\s*-+\s*)+\|$/))
			{
				if (type != 'TBL' && type != 'TBLN')
					continue;
				// scan backwards -- is this a bottom line or should it define a header above?
				match = text.length-1;
				while (match >= 0 && text[match][0] == 'TBL')
					match--;
				// Already header defined? Throw this line out:
				if (match >= 0 && text[match][0] == 'TBH')
					continue;
				// No? Then all above was a header
				if (match < 0)
					match = 0;
				else
					match++;
				if (text[match][0] == 'TBL')
				{
					while (match < text.length && text[match][0] == 'TBL')
					{
						text[match][0] = 'TBH';
						match++;
					}
					type = 'TBLN';
				}
			//	text.push (['TBLN', '--']);
			} else
			{
				type = 'TBL';
				text.push (['TBL', ln]);
			}
			continue;
		}

		// Blockquote: > ...
			// (this code was obfuscated above)
		if (ln.match(/^ {0,3}<0x005C><0x003E>/))
		{
			ln = ln.replace(/^ +/, '');
			// Multi-indent blockquote?
			match = ln.replace(/(<0x005C><0x003E> *)/g, '>').match(/^(>+)/)[1].length;

			// Empty? New line in a blockquote
			ln = ln.replace(/^(<0x005C><0x003E> *)+/, '');
			if (ln == '')
			{
				if (block != '')
				{
					text.push([type,block]);
					block = '';
				}
				continue;
			}

			// Different from previous?
			if (type == "BQ" && match != level)
			{
				if (block != '')
				{
					text.push (['BQ', block]);
					block = '';
				}
				type = "BQ";
				level = match;
				if (level > 1)
				{
					while (--match > 0)
						block += '\t';
					block += '<0x0007>';
				}
				ln = ln.replace(/^(<0x005C><0x003E> *)+/, '');
				block += ln;
				continue;
			}

			level = match;
			// Continuation of a previous line?
			if (type == "BQ" && block != '')
			{
				if (block.match(/  $/))
					block = block.substring(0, block.length-2)+"<0x000A>"+ln.replace(/^ +/, '');
				else
					block += ' '+ln;
				continue;
			}

			// Introducing new blockquote paragraph
			type = "BQ";
			block = '';
			if (level > 1)
			{
				while (--match > 0)
					block += '\t';
				block += '<0x0007>';
			}
			block += ln;
			continue;
		}

		// Indented text. This may be a new paragraph in a list, a continued
		// line of an indented item, or 'preformatted' text
		if (ln.match (/^ /))
		{
			// Blank line or plain text before?
			if (block == '' || type == 'TEXT')
			{
				// Then count actual number of spaces. Three or less is text,
				// four or more is preformatted and should be processed next
				if (ln.match(/^ {1,3}\S/))
				{
					if (block == '')
					{
						block = ln.replace (/^ */,'');
						type = 'TEXT';
					} else
					{
						if (block.match(/  $/))
							block = block.substring(0, block.length-2)+"<0x000A>"+ln.replace(/^ +/, '');
						else
							block = block+' '+ln.replace(/^ +/, '');
					}
					continue;
				}
			}

			// Continued from a previously indented item?
			// Then tack on to the end of current one
			if (block != '')
			{
				if (type == 'BULL' || type == 'NUM' || type == 'IND' || type == 'BQ')
				{
					if (block.match(/  $/))
						block = block.substring(0, block.length-2)+"<0x000A>"+ln.replace(/^ +/, '');
					else
						block += ' '+ln;
					continue;
				}
				text.push ([type, block]);
				block = '';
			}

			// Let's get a rough Level count: number of spaces divided by 4
			level = (ln.match(/^( +)/)[1].length-1) >> 2;
			if (level == 0 && (type == 'BULL' || type == 'NUM' || type == 'IND'))
			{
				// New paragraph in a bullet or number list?
				type = 'IND';
				block = ln.replace(/^ {0,4}/, '');
				continue;
			}

			// Nope. So it's preformatted -- do not gather lines
			// Take care to retain empty lines inside preformatted remain;
			// if there is a new line above, look up until there is a PRE, or something else instead.
			i = text.length;
			if (i > 0 && text[i-1][0] == 'NEWLN')
			{
				while (i > 0 && text[i-1][0] == 'NEWLN')
					i--;
				if (i > 0 && text[i-1][0] == 'PRE')
				{
					while (i < text.length)
					{
						text[i] = ['PRE',''];
						i++;
					}
				}
			}

			type = 'PRE';
			ln = ln.replace(/^ {0,4}/, '');
			text.push (['PRE', ln]);
			continue;
		}

		// Tack onto previous paragraph?
		if (block != '')
		{
			if (type == 'TEXT' || type == 'BULL' || type == 'NUM' || type == 'IND' || type == 'BQ')
			{
				if (block.match(/  $/))
					block = block.substring(0, block.length-2)+"<0x000A>"+ln.replace(/^ +/, '');
				else
					block += ' '+ln;
				continue;
			}
			text.push ([type,block]);
		}

		type = 'TEXT';
		block = ln;
	}

	return [ text, hyperlinkDefs ];
}

function accentEscape (text, hideFlag)
{
	var combo = [
		'æ', 0,  0,  0,  0,  0, 'œ', 0,  0,   0,  0,  0,  0,  0,  0,  0,  0,
		'æ', 0,  0,  0,  0,  0, 'œ', 0,  0,  'Æ', 0,  0,  0,  0,  0, 'Œ', 0,
		'á','ć','é','í', 0, 'ń','ó','ś','ú', 'Á','Ć','É','Í', 0, 'Ń','Ó','Ú',
		'à', 0, 'è','ì', 0,  0, 'ò', 0, 'ù', 'À', 0, 'È','Ì', 0,  0, 'Ò','Ù',
		'ä', 0, 'ë','ï', 0,  0, 'ö', 0, 'ü', 'Ä', 0, 'Ë','Ï', 0,  0, 'Ö','Ü',
		'ã', 0, 'ẽ','ĩ', 0, 'ñ','õ', 0, 'ũ', 'Ã', 0, 'Ẽ','Ĩ', 0, 'Ñ','Õ','Ũ',
		'â','ĉ','ê','î', 0,  0, 'ô', 0, 'û', 'Â','Ĉ','Ê','Î', 0,  0, 'Ô','Û',
		'ą','ç','ę', 0,  0, 'ņ','ǫ', 0,  0,  'Ą','Ç','Ę', 0,  0, 'Ņ','Ǫ', 0,
		 0,  0,  0,  0,  0,  0, 'ő', 0, 'ű',  0,  0,  0,  0,  0,  0, 'Ő','Ű',
		 0,  0,  0,  0, 'ł', 0,  0,  0,  0,   0,  0,  0,  0, 'Ł', 0,  0,  0,
		 0,  0,  0,  0,  0,  0,  0, 'ß',  0,  0,  0,  0,  0,  0,  0,  0,  0
		 ];
	var base = "aceilnosuACEILNOU", accent = "eE'`\"~^,#/s";
	var findReg = RegExp("<0x005C><0x005C>(["+base+"])(["+accent+"])", 'g');

	text = text.replace (/<0x005C><0x005C><0x005C><0x005C>/g, '\x85');
	if (hideFlag == true)
	{
		text = text.replace(findReg, function(match,letter,acc)
			{
				var res = base.indexOf(letter)+base.length*(accent.indexOf(acc));
				if (combo[res])
				{
					return "<0x005C><0x005C>"+letter+UCtoTag(acc.charCodeAt(0));
				}
				return match;
			}
		);
	} else
	{
		text = text.replace(findReg, function(match,letter,acc)
			{
				var res = base.indexOf(letter)+base.length*(accent.indexOf(acc));
				if (combo[res])
				{
					res = "0000"+combo[res].charCodeAt(0).toString(16);
					return "<0x"+res.substr(res.length-4)+">";
				}
				return match;
			}
		);
	}
	text = text.replace (/\x85/g, '<0x005C><0x005C>');
	return text;
}

function createText (text)
{
	tagFile = File(Folder.temp+'/md__tmp.txt');
	if (tagFile.open('w') == false)
	{
		alert ("Unable to create temporary file!");
		exit();
	}
	tagFile.encoding = "utf8";
	if (File.fs == "Windows")
		tagFile.write ("<ASCII-WIN>\n");
	else
		tagFile.write ("<ASCII-MAC>\n");
	tagFile.write (
		"<ctable:=<Blue:COLOR:RGB:Process:0,0,1>><dps:"+tagset['text']+"=<psb:6><psa:6>>\n" +
		"<dps:"+tagset['h1']+"=<BasedOn:"+tagset['text']+"><pKeepWithNext:1><ct:Bold><cs:20><psb:24><psa:12>>\n" +
		"<dps:"+tagset['h2']+"=<BasedOn:"+tagset['text']+"><pKeepWithNext:1><ct:Bold><cs:16><psb:12><psa:12>>\n" +
		"<dps:"+tagset['h3']+"=<BasedOn:"+tagset['text']+"><pKeepWithNext:1><ct:Italic><cs:14><psb:12><psa:6>>\n" +
		"<dps:"+tagset['h4']+"=<BasedOn:"+tagset['text']+"><pKeepWithNext:1><ct:Italic><cs:12><psb:12><psa:0>>\n" +
		"<dps:"+tagset['h5']+"=<BasedOn:"+tagset['text']+"><pKeepWithNext:1><cu:1><ct:Italic><cs:12><psb:6><psa:0>>\n" +
		"<dps:"+tagset['h6']+"=<BasedOn:"+tagset['text']+"><pKeepWithNext:1><cu:1><cs:12><psb:6><psa:0>>\n" +
		"<dps:"+tagset['bullet']+"=<BasedOn:"+tagset['text']+"><bnlt:Bullet><pli:36><pfli:-36><psb:3><psa:3>>\n" +
		"<dps:"+tagset['numbered']+"=<BasedOn:"+tagset['text']+"><bnlt:Numbered><pli:36><pfli:-36><psb:3><psa:3>>\n" +
		"<dps:"+tagset['indent']+"=<BasedOn:"+tagset['text']+"><pli:36><psb:3><psa:3>>\n" +
		"<dps:"+tagset['quote']+"=<BasedOn:"+tagset['text']+"><pli:36>>\n" +
		"<dps:"+tagset['pre']+"=<BasedOn:"+tagset['text']+"><cf:Courier New><psb:0><psa:0>>\n" +
		"<dps:"+tagset['th']+"=<BasedOn:"+tagset['text']+"><ct:Bold>>\n" +
		"<dps:"+tagset['td']+"=<BasedOn:"+tagset['text']+">>\n" +
		"<dps:"+tagset['hrule']+"=<BasedOn:"+tagset['text']+"><pRuleBelowOn:1><pRuleBelowOffset:-3>>\n" +
		"<dcs:"+tagset['em']+"=<ct:Italic>>\n" +
		"<dcs:"+tagset['strong']+"=<ct:Bold>>\n" +
		"<dcs:"+tagset['strongem']+"=<ct:Bold Italic>>\n" +
		"<dcs:"+tagset['code']+"=<cf:Courier New>>\n" +
		"<dcs:"+tagset['hyperlink']+"=<cu:1><cc:Blue>>\n"+
		"<dtbls:"+tagset['table']+'=>\n' +
		"<des:"+tagset['header']+'=<tcsps:'+tagset['th']+'>>\n' +
		"<des:"+tagset['body']+'=<tcsps:'+tagset['td']+'>>\n'
	);
	tagFile.write (text.join("\n")+'\n');
	tagFile.close();
	// Import Tagged file;
	// temporarily switch off smart hyphens, as we supply our own!
	smartQuotes = app.taggedTextImportPreferences.useTypographersQuotes;
	app.taggedTextImportPreferences.useTypographersQuotes = false;
	app.place (File(Folder.temp+'/md__tmp.txt'), false);
//	app.place (File(Folder.myDocuments+'/md__tmp.txt'), false);
	app.taggedTextImportPreferences.useTypographersQuotes = smartQuotes;
}

function UCtoTag(a)
{
	a = "0000"+a.toString(16);
	return "<0x"+a.substring(a.length-4)+">";
}

function UCtoHiddenTag(a)
{
	a = "0000"+a.toString(16);
	return "\xAE0x"+a.substring(a.length-4)+"\xAF";
}
