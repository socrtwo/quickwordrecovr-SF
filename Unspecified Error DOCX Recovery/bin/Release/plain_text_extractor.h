#ifndef DOCTOTEXT_PLAIN_TEXT_EXTRACTOR_H
#define DOCTOTEXT_PLAIN_TEXT_EXTRACTOR_H

#include "formatting_style.h"
#include <string>

class Metadata;

/**
	Extracts plain text from documents. In addition it can be used to extract metadata and comments (annotations).
	Example of usage (extracting plain text):
	\code
	PlainTextExtractor extractor;
	std::string text;
	if (extractor.processFile("example.doc", text))
		std::cout << text << std::endl;
	else
		std::cerr << "Error." << std::endl;
	\endcode
	Example of usage (extracting metadata):
	\code
	PlainTextExtractor extractor;
	Metadata meta;
	if (extractor.extractMetadata("example.doc", meta))
		std::cout << meta.author << std::endl;
	else
		std::cerr << "Error." << std::endl;
	\endcode
**/
class PlainTextExtractor
{
	private:
		struct Implementation;
		Implementation* impl;

	public:

		/**
			Enumerates all supported document formats. \c PARSER_AUTO means unknown format that should be
			determined.
		**/
		enum ParserType { PARSER_AUTO, PARSER_RTF, PARSER_ODF_OOXML, PARSER_XLS, PARSER_DOC, PARSER_PPT, PARSER_HTML };

		/**
			The constructor.
			\param parser_type restricts parser to specified document format. If set to \c PARSER_AUTO the parser
				will work with all supported documents formats.
		**/
		PlainTextExtractor(ParserType parser_type = PARSER_AUTO);

		~PlainTextExtractor();

		/**
			Enables or disables verbose logging. Verbose logging is disabled by default.
			If verbose logging is disabled only important messages and errors are logged.
			If verbose logging is enabled all messages and errors are logged.
			\warning Verbose logging can produce a lot of text, especially if the library was compiled in debug
			mode.
			\param verbose if \c true verbose logging will be enabled. If \c false verbose logging will be
			disabled.
			\see setLogStream
		**/
		void setVerboseLogging(bool verbose);

		/**
			Assign an output stream that will be used for logging messages and errors.
			It can be used to capture logs to a file, string or show them in dialog.
			\c std::cerr stream is used by default.
			\param log_stream the stream that will be used for logging
			\see setVerboseLogging
		**/
		void setLogStream(std::ostream& log_stream);

		/**
			Sets how tables, lists and urls should be formatted in plain text produced by them
			parser.
			\param style instance of structure \c FormattingStyle that specifies formatting style.
			\see FormattingStyle
		**/
		void setFormattingStyle(const FormattingStyle& style);

		void setXmlParseMode(XmlParseMode mode);

		/**
			Tries to determine document format by file name extension.
			\warning Some applications save CSV documents with "xls" extension, RTF documents with "doc"
			extension or HTML documents with "xls" or "doc" extension. In such a situation this simple test
			will fail.
			\param file_name file name or full path to file.
			\return value of \c ParserType type representing determined document format or PARSER_AUTO if
			document format cannot be determined.
			\see ParserType parserTypeByFileContent
		**/
		ParserType parserTypeByFileExtension(const std::string& file_name);

		/**
			\overload
		**/
		ParserType parserTypeByFileExtension(const char* file_name);

		/**
			Tries to determine document format by file content.
			\param file_name full path to file containing document.
			\param reference to variable of \c ParserType type that will contain determined document format
			or PARSER_AUTO if document format cannot be determined.
			\return \c true if document was processed successfully, \c false otherwise.
			\see ParserType parserTypeByFileExtension
		**/
		bool parserTypeByFileContent(const std::string& file_name, ParserType& parser_type);

		/**
			\overload
		**/
		bool parserTypeByFileContent(const char* file_name, ParserType& parser_type);

		/**
			Parses specified document and extracts plain text.
			\param file_name full path to file containing document.
			\param text reference to object of \c std::string class that will contain produced plain text.
			\return \c true if document was processed successfully, \c false otherwise.
			\see ParserType setFormattingStyle
		**/
		bool processFile(const std::string& file_name, std::string& text);

		/**
			\overload
			\param text reference to pointer that will point to produced plain text in form of null-terminated
				array of chars. The caller is responsible for deleting the buffer using \c delete[] operator.
		**/
		bool processFile(const char* file_name, char*& text);

		/**
			Parses specified document and extracts plain text.
			\param parser_type restricts parser to specified document format. If set to \c PARSER_AUTO the parser
				will work with all supported documents formats. This argument override parser type set for
				the object.
			\param fallback if \c true parser will try to detect document format if parsing of document format
				specified in \c parser_type argument fails. This parameter is ignored if \c parser_type is
				set to \c PARSER_AUTO.
			\param file_name full path to file containing document.
			\param text reference to object of \c std::string class that will contain produced plain text.
			\return \c true if document was processed successfully, \c false otherwise.
			\see ParserType setFormattingStyle
		**/
		bool processFile(ParserType parser_type, bool fallback, const std::string& file_name, std::string& text);

		/**
			\overload
			\param text reference to pointer that will point to produced plain text in form of null-terminated
				buffer. The caller is responsible for deleting the buffer using \c delete[] operator.
		**/
		bool processFile(ParserType parser_type, bool fallback, const char* file_name, char*& text);

		/**
			Parses specified document and extracts metadata (author, creation time, etc).
			\param file_name full path to file containing document.
			\param metadata reference to object of \c Metadata class that will contain extracted information.
			\return \c true if document was processed successfully, \c false otherwise.
			\see ParserType Metadata
		**/
		bool extractMetadata(const std::string& file_name, Metadata& metadata);

		/**
			\overload
		**/
		bool extractMetadata(const char* file_name, Metadata& metadata);

		/**
			Parses specified document and extracts metadata (author, creation time, etc).
			\param parser_type restricts parser to specified document format. If set to \c PARSER_AUTO the parser
				will work with all supported documents formats. This argument override parser type set for
				the object.
			\param fallback if \c true parser will try to detect document format if parsing of document format
				specified in \c parser_type argument fails. This parameter is ignored if \c parser_type is
				set to \c PARSER_AUTO.
			\param file_name full path to file containing document.
			\param metadata reference to object of \c Metadata class that will contain extracted information.
			\return \c true if document was processed successfully, \c false otherwise.
			\see ParserType Metadata
		**/
		bool extractMetadata(ParserType parser_type, bool fallback, const std::string& file_name, Metadata& metadata);

		/**
			\overload
		**/
		bool extractMetadata(ParserType parser_type, bool fallback, const char* file_name, Metadata& metadata);
};

#endif
