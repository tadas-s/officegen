//
// officegen: All the code to generate DOCX files.
//
// Please refer to README.md for this module's documentations.
//
// NOTE:
// - Before changing this code please refer to the hacking the code section on README.md.
//
// Copyright (c) 2013 Ziv Barber;
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// 'Software'), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
// IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
// CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
// TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

var baseobj = require("./basicgen.js");
var msdoc = require("./msofficegen.js");

var path = require('path');

var fast_image_size = require('fast-image-size');

if ( !String.prototype.encodeHTML ) {
	String.prototype.encodeHTML = function () {
		return this.replace(/&/g, '&amp;')
			.replace(/</g, '&lt;')
			.replace(/>/g, '&gt;')
			.replace(/"/g, '&quot;');
	};
}

///
/// @brief Extend officegen object with DOCX support.
///
/// This method extending the given officegen object to create DOCX document.
///
/// @param[in] genobj The object to extend.
/// @param[in] new_type The type of object to create.
/// @param[in] options The object's options.
/// @param[in] gen_private Access to the internals of this object.
/// @param[in] type_info Additional information about this type.
///
function makeDocx ( genobj, new_type, options, gen_private, type_info ) {
	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxApp ( data ) {
		var userName = genobj.options.creator || 'officegen';
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template>Normal.dotm</Template><TotalTime>1</TotalTime><Pages>1</Pages><Words>0</Words><Characters>0</Characters><Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><Lines>1</Lines><Paragraphs>1</Paragraphs><ScaleCrop>false</ScaleCrop><Company>' + userName + '</Company><LinksUpToDate>false</LinksUpToDate><CharactersWithSpaces>0</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>';
		return outString;
	}

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] data Ignored by this callback function.
	/// @return Text string.
	///
	function cbMakeDocxDocument ( data ) {
		var outString = gen_private.plugs.type.msoffice.cbMakeMsOfficeBasicXml ( data ) + '<w:document xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><w:body>';
		var objs_list = data.data;

		// Work on all the stored paragraphs inside this document:
		for ( var i = 0, total_size = objs_list.length; i < total_size; i++ ) {
			outString += '<w:p w:rsidR="00A77427" w:rsidRDefault="007F1D13">';
			var pPrData = '';

			if ( objs_list[i].options ) {
				if ( objs_list[i].options.align ) {
					switch ( objs_list[i].options.align ) {
						case 'center':
							pPrData += '<w:jc w:val="center"/>';
							break;

						case 'right':
							pPrData += '<w:jc w:val="right"/>';
							break;

						case 'justify':
							pPrData += '<w:jc w:val="both"/>';
							break;
					} // End of switch.
				} // Endif.

				if ( objs_list[i].options.list_type ) {
					pPrData += '<w:pStyle w:val="Normal"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + objs_list[i].options.list_type + '"/></w:numPr>';
				} // Endif.

        /*
         * Paragraph spacing options
         *
         * More details: http://officeopenxml.com/WPspacing.php
         */
        if ( objs_list[i].options.spacing ) {
          var spacing_options = [];

          if ( typeof objs_list[i].options.spacing.before !== 'undefined' ) {
            spacing_options.push('w:before="' + parseInt(objs_list[i].options.spacing.before, 10) + '"');
          }

          if ( typeof objs_list[i].options.spacing.after !== 'undefined' ) {
            spacing_options.push('w:after="' + parseInt(objs_list[i].options.spacing.after, 10) + '"');
          }

          if ( typeof objs_list[i].options.spacing.line !== 'undefined' ) {
            spacing_options.push('w:line="' + Math.round(240 * parseFloat(objs_list[i].options.spacing.line, 10)) + '"');
          }

          if ( spacing_options.length ) {
            pPrData += '<w:spacing ' + spacing_options.join(' ') + ' w:lineRule="auto"/>';
          }
        } // Endif

        /*
         * Paragraph indentation options
         *
         * More details: http://officeopenxml.com/WPindentation.php
         */
        if ( objs_list[i].options.indentation ) {
          var indentation_options = [];

          if ( typeof objs_list[i].options.indentation.left !== 'undefined' ) {
            indentation_options.push('w:left="' + parseInt(objs_list[i].options.indentation.left, 10) + '"');
          }

          if ( typeof objs_list[i].options.indentation.right !== 'undefined' ) {
            indentation_options.push('w:right="' + parseInt(objs_list[i].options.indentation.right, 10) + '"');
          }

          if ( typeof objs_list[i].options.indentation.hanging !== 'undefined' ) {
            indentation_options.push('w:hanging="' + parseInt(objs_list[i].options.indentation.hanging, 10) + '"');
          }

          if ( indentation_options.length ) {
            pPrData += '<w:ind ' + indentation_options.join(' ') + '/>';
          }
        } // Endif
			} // Endif.

			if ( pPrData ) {
				outString += '<w:pPr>' + pPrData + '</w:pPr>';
			} // Endif.

			// Work on all the objects in the document:
			for ( var j = 0, total_size_j = objs_list[i].data.length; j < total_size_j; j++ ) {
				if ( objs_list[i].data[j] ) {
					var rExtra = '';
					var tExtra = '';
					var rPrData = '';

					if ( objs_list[i].data[j].options ) {
						if ( objs_list[i].data[j].options.color ) {
							rPrData += '<w:color w:val="' + objs_list[i].data[j].options.color + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.back ) {
							rPrData += '<w:shd w:val="clear" w:color="auto" w:fill="' + objs_list[i].data[j].options.back + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.bold ) {
							rPrData += '<w:b/><w:bCs/>';
						} // Endif.

						if ( objs_list[i].data[j].options.italic ) {
							rPrData += '<w:i/><w:iCs/>';
						} // Endif.

						if ( objs_list[i].data[j].options.underline ) {
							rPrData += '<w:u w:val="single"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.font_face ) {
							rPrData += '<w:rFonts w:ascii="' + objs_list[i].data[j].options.font_face + '" w:hAnsi="' + objs_list[i].data[j].options.font_face + '" w:cs="' + objs_list[i].data[j].options.font_face + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.font_size ) {
              var fontSizeInHalfPoints = 2 * objs_list[i].data[j].options.font_size;
							rPrData += '<w:sz w:val="' + fontSizeInHalfPoints + '"/><w:szCs w:val="' + fontSizeInHalfPoints + '"/>';
						} // Endif.

						if ( objs_list[i].data[j].options.border ) {
							switch ( objs_list[i].data[j].options.border )
							{
								case 'single':
								case true:
									rPrData += '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
									break;
							} // End of switch.
						} // Endif.
					} // Endif.

					if ( objs_list[i].data[j].text ) {
            if ( (objs_list[i].data[j].text[0] == ' ') || (objs_list[i].data[j].text[objs_list[i].data[j].text.length - 1] == ' ') ) {
							tExtra += ' xml:space="preserve"';
						} // Endif.

						outString += '<w:r' + rExtra + '>';

						if ( rPrData ) {
							outString += '<w:rPr>' + rPrData + '</w:rPr>';
						} // Endif.

						outString += '<w:t' + tExtra + '>' + objs_list[i].data[j].text.encodeHTML () + '</w:t></w:r>';

					} else if ( objs_list[i].data[j].page_break ) {
						outString += '<w:r><w:br w:type="page"/></w:r>';

					} else if ( objs_list[i].data[j].line_break ) {
						outString += '<w:r><w:br/></w:r>';

					} else if ( objs_list[i].data[j].image ) {
						outString += '<w:r' + rExtra + '>';

						rPrData += '<w:noProof/>';

						if ( rPrData ) {
							outString += '<w:rPr>' + rPrData + '</w:rPr>';
						} // Endif.

						//914400L / 96DPI
						var pixelToEmu = 9525;

						outString += '<w:drawing>';
						outString += '<wp:inline distT="0" distB="0" distL="0" distR="0">';
						outString += '<wp:extent cx="' + (objs_list[i].data[j].options.cx * pixelToEmu) + '" cy="' + (objs_list[i].data[j].options.cy * pixelToEmu) + '"/>';
						outString += '<wp:effectExtent l="19050" t="0" r="9525" b="0"/>';
						outString += '<wp:docPr id="' + (objs_list[i].data[j].image_id + 1) + '" name="Picture ' + objs_list[i].data[j].image_id + '" descr="Picture ' + objs_list[i].data[j].image_id + '"/>';
						outString += '<wp:cNvGraphicFramePr>';
						outString += '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>';
						outString += '</wp:cNvGraphicFramePr>';
						outString += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
						outString += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">';
						outString += '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">';
						outString += '<pic:nvPicPr>';
						outString += '<pic:cNvPr id="0" name="Picture ' + objs_list[i].data[j].image_id + '"/>';
						outString += '<pic:cNvPicPr/>';
						outString += '</pic:nvPicPr>';
						outString += '<pic:blipFill>';
						outString += '<a:blip r:embed="rId' + objs_list[i].data[j].rel_id + '" cstate="print"/>';
						outString += '<a:stretch>';
						outString += '<a:fillRect/>';
						outString += '</a:stretch>';
						outString += '</pic:blipFill>';
						outString += '<pic:spPr>';
						outString += '<a:xfrm>';
						outString += '<a:off x="0" y="0"/>';
						outString += '<a:ext cx="' + (objs_list[i].data[j].options.cx * pixelToEmu) + '" cy="' + (objs_list[i].data[j].options.cy * pixelToEmu) + '"/>';
						outString += '</a:xfrm>';
						outString += '<a:prstGeom prst="rect">';
						outString += '<a:avLst/>';
						outString += '</a:prstGeom>';
						outString += '</pic:spPr>';
						outString += '</pic:pic>';
						outString += '</a:graphicData>';
						outString += '</a:graphic>';
						outString += '</wp:inline>';
						outString += '</w:drawing>';

						outString += '</w:r>';
					} // Endif.
				} // Endif.
			} // Endif.

			outString += '</w:p>';
		} // End of for loop.

		outString += '<w:p w:rsidR="00A02F19" w:rsidRDefault="00A02F19"/><w:sectPr w:rsidR="00A02F19" w:rsidSect="00A02F19"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>';
		return outString;
	}

	// Prepare genobj for MS-Office:
	msdoc.makemsdoc ( genobj, new_type, options, gen_private, type_info );
	gen_private.plugs.type.msoffice.makeOfficeGenerator ( 'word', 'document', {} );

	genobj.on ( 'clearData', function () {
		genobj.data.length = 0;
	});

	gen_private.plugs.type.msoffice.addInfoType ( 'dc:title', '', 'title', 'setDocTitle' );
	gen_private.plugs.type.msoffice.addInfoType ( 'dc:subject', '', 'subject', 'setDocSubject' );
	gen_private.plugs.type.msoffice.addInfoType ( 'cp:keywords', '', 'keywords', 'setDocKeywords' );
	gen_private.plugs.type.msoffice.addInfoType ( 'dc:description', '', 'description', 'setDescription' );

	gen_private.type.msoffice.files_list.push (
		{
			name: '/word/settings.xml',
			type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
			clear: 'type'
		},
		{
			name: '/word/fontTable.xml',
			type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
			clear: 'type'
		},
		{
			name: '/word/webSettings.xml',
			type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
			clear: 'type'
		},
		{
			name: '/word/styles.xml',
			type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
			clear: 'type'
		},
    {
      name: '/word/numbering.xml',
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
      clear: 'type'
    },
		{
			name: '/word/document.xml',
			type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
			clear: 'type'
		}
	);

	gen_private.type.msoffice.rels_app.push (
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
			target: 'styles.xml',
			clear: 'type'
		},
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
			target: 'settings.xml',
			clear: 'type'
		},
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
			target: 'webSettings.xml',
			clear: 'type'
		},
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
			target: 'fontTable.xml',
			clear: 'type'
		},
		{
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
			target: 'theme/theme1.xml',
			clear: 'type'
		},
    {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
      target: 'numbering.xml',
      clear: 'type'
    }
	);

	genobj.data = []; // All the data will be placed here.

	gen_private.plugs.intAddAnyResourceToParse ( 'docProps\\app.xml', 'buffer', null, cbMakeDocxApp, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'word\\fontTable.xml', 'file', path.join(__dirname, 'tmpl/docx/fontTable.xml'), null, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'word\\settings.xml', 'file', path.join(__dirname, 'tmpl/docx/settings.xml'), null, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'word\\webSettings.xml', 'file', path.join(__dirname, 'tmpl/docx/webSettings.xml'), null, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'word\\styles.xml', 'file', path.join(__dirname, 'tmpl/docx/styles.xml'), null, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'word\\numbering.xml', 'file', path.join(__dirname, 'tmpl/docx/numbering.xml'), null, true );
	gen_private.plugs.intAddAnyResourceToParse ( 'word\\document.xml', 'buffer', genobj, cbMakeDocxDocument, true );

	gen_private.plugs.intAddAnyResourceToParse ( 'word\\_rels\\document.xml.rels', 'buffer', gen_private.type.msoffice.rels_app, gen_private.plugs.type.msoffice.cbMakeRels, true );

	// ----- API for Word documents: -----

	///
	/// @brief Create a new paragraph.
	///
	/// ???.
	///
	/// @param[in] options Default options for all the objects inside this paragraph.
	///
	genobj.createP = function ( options ) {
		var newP = {};

		newP.data = [];
		newP.options = options || {};

		///
		/// @brief Insert text inside this paragraph.
		///
		/// ???.
		///
		/// @param[in] text_msg The text message itself.
		/// @param[in] opt ???.
		/// @param[in] flag_data ???.
		///
		newP.addText = function ( text_msg, opt, flag_data ) {
			newP.data[newP.data.length] = { text: text_msg, options: opt || {}, ext_data: flag_data };
		};

		///
		/// @brief Insert a line break inside this paragraph.
		///
		/// ???.
		///
		///
		newP.addLineBreak = function () {
      newP.data[newP.data.length] = { 'line_break': true };
		};

		///
		/// @brief Insert an image into the current paragraph.
		///
		/// ???.
		///
		/// @param[in] image_path The image file to add.
		/// @param[in] opt Additional options (cx, cy).
		/// @param[in] image_format_type ???.
		///
		newP.addImage = function ( image_path, opt, image_format_type ) {
			var image_type = (typeof image_format_type == 'string') ? image_format_type : 'png';
			var defWidth = 320;
			var defHeight = 200;

			if ( typeof image_path == 'string' ) {
				var ret_data = fast_image_size ( image_path );
				if ( ret_data.type == 'unknown' ) {
					var image_ext = path.extname ( image_path );

					switch ( image_ext ) {
						case '.bmp':
							image_type = 'bmp';
							break;

						case '.gif':
							image_type = 'gif';
							break;

						case '.jpg':
						case '.jpeg':
							image_type = 'jpeg';
							break;

						case '.emf':
							image_type = 'emf';
							break;

						case '.tiff':
							image_type = 'tiff';
							break;
					} // End of switch.

				} else {
					if ( ret_data.width ) {
						defWidth = ret_data.width;
					} // Endif.

					if ( ret_data.height ) {
						defHeight = ret_data.height;
					} // Endif.

					image_type = ret_data.type;
					if ( image_type == 'jpg' ) {
						image_type = 'jpeg';
					} // Endif.
				} // Endif.
			} // Endif.

			var objNum = newP.data.length;
			newP.data[objNum] = { image: image_path, options: opt || {} };

			if ( !newP.data[objNum].options.cx && defWidth ) {
				newP.data[objNum].options.cx = defWidth;
			} // Endif.

			if ( !newP.data[objNum].options.cy && defHeight ) {
				newP.data[objNum].options.cy = defHeight;
			} // Endif.

			var image_id = gen_private.type.msoffice.src_files_list.indexOf ( image_path );
			var image_rel_id = -1;

			if ( image_id >= 0 ) {
				for ( var j = 0, total_size_j = gen_private.type.msoffice.rels_app.length; j < total_size_j; j++ ) {
					if ( gen_private.type.msoffice.rels_app[j].target == ('media/image' + (image_id + 1) + '.' + image_type) ) {
						image_rel_id = j + 1;
					} // Endif.
				} // Endif.

			} else
			{
				image_id = gen_private.type.msoffice.src_files_list.length;
				gen_private.type.msoffice.src_files_list[image_id] = image_path;
				gen_private.plugs.intAddAnyResourceToParse ( 'word\\media\\image' + (image_id + 1) + '.' + image_type, (typeof image_path == 'string') ? 'file' : 'stream', image_path, null, false );
			} // Endif.

			if ( image_rel_id == -1 ) {
				image_rel_id = gen_private.type.msoffice.rels_app.length + 1;

				gen_private.type.msoffice.rels_app.push (
					{
						type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
						target: 'media/image' + (image_id + 1) + '.' + image_type,
						clear: 'data'
					}
				);
			} // Endif.

			newP.data[objNum].image_id = image_id;
			newP.data[objNum].rel_id = image_rel_id;
		};

		genobj.data[genobj.data.length] = newP;
		return newP;
	};

	///
	/// @brief ???.
	///
	/// ???.
	///
	/// @param[in] options ???.
	///
	genobj.createListOfDots = function ( options ) {
		var newP = genobj.createP ( options );

    //TODO: have these coincide with the numbers.xml
		newP.options.list_type = '1';

		return newP;
	};

	///
	/// @brief Create a list of numbers based paragraph.
	///
	/// ???.
	///
	/// @param[in] options ???.
	///
	genobj.createListOfNumbers = function ( options ) {
		var newP = genobj.createP ( options );

		newP.options.list_type = '1';

		return newP;
	};

	///
	/// @brief Add a page break.
	///
	/// This method add a page break to the current Word document.
	///
	genobj.putPageBreak = function () {
		var newP = {};

		newP.data = [ { 'page_break': true } ];

		genobj.data[genobj.data.length] = newP;
		return newP;
	};

	///
	/// @brief Add a page break.
	///
	/// This method add a page break to the current Word document.
	///
	genobj.addPageBreak = function () {
		var newP = {};

		newP.data = [ { 'page_break': true } ];

		genobj.data[genobj.data.length] = newP;
		return newP;
	};
}

baseobj.plugins.registerDocType ( 'docx', makeDocx, {}, baseobj.docType.TEXT, "Microsoft Word Document" );

