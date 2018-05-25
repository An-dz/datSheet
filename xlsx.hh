#include <vector>      // vector
#include <zip.hpp>     // libzip++
#include <pugixml.hpp> // pugixml

/**
 * Parser for Office Open XML xlsx documents
 */
class XLSX
{
	/** pointer to loaded spreadsheet xlsx file */
	libzip::archive* sheet;
	/** XML DOM node that holds the strings in the xlsx */
	pugi::xml_node strings_xml;
	/** structure that holds important sheet data
	 *
	 * @note sheet id and name are stored in the workbook xml
	 * but the path is only stored in the .rels file of the
	 * workbook.
	 */
	struct sheet_t {
		std::string id;
		std::string name;
		std::string path;
	};
	/** vector that holds info about each sheet */
	std::vector<sheet_t> sheets_v;

	// Get a DOM object of an XML inside the zip
	void xml_open(const std::string& filename, pugi::xml_document& doc);
	// Create the dat files
	void createDat(const pugi::xml_node& node, const unsigned char sheet_nr, std::string*const dat_parameters, std::string& last_filename);
	// Write the dat file on disk
	const unsigned int writeDat(const std::string& filename, const std::string& dat_stream, const bool append);

public:
	// Open an xlsx file
	XLSX(const std::string& filename);
	// Destructor
	~XLSX();
	// Parse an xlsx file
	void parse();
};
