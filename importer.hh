#include <vector>      // vector
#include <zip.hpp>     // libzip++
#include <pugixml.hpp> // pugixml

#define VERSION "1.2.0"

/**
 * Importer from a directory tree to a valid Open Office XML xlsx
 */
class Importer
{
	/** pointer to loaded spreadsheet xlsx file */
	libzip::archive* sheet;
	/** xlsx sharedStrings file */
	std::vector<std::string> sharedStrings;
	/** name for the worksheets */
	std::vector<std::string> worksheets;
	unsigned int root_dir_size;

	// Adds the xml declaration header
	void addXMLdeclaration(pugi::xml_document& doc);
	// Converts a string from whatever encoding to UTF-8
	bool convertToUTF8(const std::string& input, std::stringstream& output);
	// Add the sheet file in the zip
	void createSheet(const std::vector<std::string>& dats, const std::string& dir, const unsigned int index);
	// Searches a vector and add value if not found
	unsigned int findInVectorOrAdd(std::vector<std::string>& str_v, const std::string& value);
	// Iterate over directory to find results
	void readDir(const std::string& dir_name, unsigned int& index);
public:
	// Create an xlsx file
	Importer(const std::string& filename);
	// Destructor to remove sheet from memory
	~Importer();
	// Start importing
	void import(const std::string& root_dir);
};
