#include <iostream> // cout, cerr, clog
#include <sstream>  // ostringstream
#include <fstream>  // ofstream
#include "xlsx.hh"

/**
 * @brief Open an xlsx file
 *
 * An xlsx file is a normal zip file with multiple xmls inside.
 * This will open the zip for reding the files inside.
 *
 * @param filename Name of the spreadsheet file
 */
XLSX::XLSX(const std::string& filename)
{
	try {
		// open as read-only
		this->sheet = new libzip::archive(filename, ZIP_RDONLY);
		return;
	}
	catch (const std::runtime_error& e) {
		std::ostringstream err_msg;
		err_msg << "ZIP" << errno << ":" << e.what() << ": " << filename;
		// send to main
		throw std::runtime_error(err_msg.str());
	}
}

/**
 * @brief Destroy object
 *
 * Destroys the object removing the loaded xlsx from memory
 */
XLSX::~XLSX()
{
	delete sheet;
}

/**
 * @brief Parse an xlsx file
 *
 * This function makes the heavy work of parsing the file.
 */
void XLSX::parse()
{
	// read root .rels file, contains information about file structure
	pugi::xml_document doc;
	xml_open("_rels/.rels", doc);

	// find where's the root workbook document
	const std::string workbook_path = doc.child("Relationships").find_child_by_attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument").attribute("Target").value();

	xml_open(workbook_path, doc);

	// get sheets id and name and put in vector
	for (const pugi::xml_node sheet: doc.child("workbook").child("sheets").children()) {
		sheet_t sheet_info = {
			id: sheet.attribute("r:id").value(),
			name: sheet.attribute("name").value()
		};
		sheets_v.push_back(sheet_info);
	}

	// open relations file inside the workbook dir to get where are the sheets and where are the strings stored
	const std::string::size_type pos = workbook_path.rfind("/");
	const std::string spreadsheet_path = workbook_path.substr(0, pos);
	const std::string workbook_rels = spreadsheet_path + "/_rels" + workbook_path.substr(pos) + ".rels";

	xml_open(workbook_rels, doc);

	// get the relative location of each sheet
	for (unsigned int i = 0; i < sheets_v.size(); ++i) {
		sheets_v[i].path = spreadsheet_path + "/" + doc.child("Relationships").find_child_by_attribute("Id", sheets_v[i].id.c_str()).attribute("Target").value();
	}

	// get where are the strings stored
	const std::string strings_file = spreadsheet_path + "/" + doc.child("Relationships").find_child_by_attribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings").attribute("Target").value();

	xml_open(strings_file, doc);
	strings_xml = doc.child("sst");

	// open the sheets and work on them
	for (unsigned int i = 0; i < sheets_v.size(); ++i) {
		pugi::xml_document sheet_doc;
		xml_open(sheets_v[i].path, sheet_doc);
		const pugi::xml_node sheetData = sheet_doc.child("worksheet").child("sheetData");

		// array that will contain the dat parameters names
		std::string dat_parameters[255];
		std::string last_filename;

		// paramater names are in the first row, let's cache them
		createDat(sheetData.find_child_by_attribute("r", "1"), i, dat_parameters, last_filename);

		// generate the dats
		for (const pugi::xml_node row: sheetData.children()) {
			createDat(row, i, dat_parameters, last_filename);
		}
	}
}

/**
 * @brief Get a DOM object of an XML inside the zip
 *
 * Opens an XML that is inside the container, reads it and places
 * it in the `doc` DOM node object for easy manipulation.
 *
 * @param filename Filename relative to the container root
 * @param doc XML DOM node where the loaded data will be placed
 */
void XLSX::xml_open(const std::string& filename, pugi::xml_document& doc)
{
	try {
		const std::string xml_data = sheet->open(filename).read(sheet->stat(filename).size);
		const pugi::xml_parse_result result = doc.load_string(xml_data.c_str());

		if (!result) {
			std::ostringstream err_msg;
			err_msg << "XML" << result.status << ":" << result.description();
			// send to main
			throw std::runtime_error(err_msg.str());
		}

		return;
	}
	catch (const std::runtime_error& e) {
		std::ostringstream err_msg;
		err_msg << "ZIP" << errno << ":" << e.what() << ": " << filename;
		// send to main
		throw std::runtime_error(err_msg.str());
	}
}

/**
 * @brief Create the dat files
 *
 * Reads the passed row data checking if cells are valid and
 * builds the dat file in a stream to later write it at once.
 *
 * @note The file name can be set anywhere in the row so we
 * can only write the dat once we have the name of the file.
 *
 * @param row_node XML node of a single XLSX row
 * @param sheet_nr Internal number of the sheet where the
 * data belongs to
 * @param dat_parameters Pointer to the array that contains
 * the paramaters' names
 * @param last_filename Pointer to a string that holds the
 * filename of the previously created file to check if user
 * wants to append to the same dat
 */
void XLSX::createDat(const pugi::xml_node& row_node, const unsigned char sheet_nr, std::string *const dat_parameters, std::string& last_filename)
{
	const std::string row_number = row_node.attribute("r").value();

	if (row_number != "1" && !row_node.find_child_by_attribute("r", ("A" + row_number).c_str())) {
		return;
	}

	std::string filename;
	std::ostringstream dat_stream;

	for (const pugi::xml_node cell: row_node.children()) {
		const std::string cell_pos = cell.attribute("r").value();
		const std::string type = cell.attribute("t").value();
		std::string value = cell.child_value("v");

		// not a number
		if (type != "" && type != "n") {
			// string
			if (type == "s") {
				value = std::to_string(stoi(value) + 1);
				value = strings_xml.select_node(("si[" + value + "]").c_str()).node().child_value("t");
			}
			// boolean
			else if (type == "b") {
				value = stoi(value) ? "true" : "false";
			}
			else if (type == "inlineStr") {
				value = cell.child_value("is");
			}
			else {
				std::clog << sheets_v[sheet_nr].name << "(" << cell_pos << ") : Wrong type warning DATAT" << type << ":Data type at " << cell_pos << " is not of expected type!\n\tExpected types: Number, Boolean, String, InlineString\n";
			}
		}

		// get column letter code and transform in a number
		const std::string column_str = cell_pos.substr(0, cell_pos.rfind(row_number));
		const unsigned int column_size = column_str.size() - 1;
		unsigned int column = 0;
		for (unsigned int i = 0; i <= column_size; ++i) {
			column = column_str[i] - 'A' + (column_size - i) * 25;
		}

		// if we are dealing with the first row we save the parameters for later use
		if (row_number == "1") {
			*(dat_parameters + column) = value;
		}
		// if not build the dat
		else if ((dat_parameters + column)->size() > 0) {
			bool is_filename = (*(dat_parameters + column) == "filename");

			if (!is_filename) {
				dat_stream << *(dat_parameters + column) << ((dat_parameters + column)->front() == '#' ? " " : "=") << value << std::endl;
			}

			if (is_filename || (filename.empty() && *(dat_parameters + column) == "name")) {
				filename = value;
			}
		}

		// std::cout << cell_pos << " | " << type << " | " << value << std::endl;
	}

	// std::cout << row_number << std::endl;

	// don't generate dat for the first row which is reserved for the dat parameters
	if (row_number != "1") {
		// std::cout << dat_stream.str() << std::endl;
		// std::cout << (filename == last_filename) << " : " << filename << " > " << last_filename << std::endl;

		std::string sheet_name = sheets_v[sheet_nr].name;

		std::string::size_type pos = sheet_name.find(";");
		while (pos != std::string::npos) {
			sheet_name = sheet_name.replace(pos, 1, "/");
			pos = sheet_name.find(";");
		}

		switch (writeDat(sheet_name + "/" + filename, dat_stream.str(), filename == last_filename)) {
			default:
				break;
			case 1:
				std::clog << sheets_v[sheet_nr].name << "(" << row_number << ") : No name warning FDATOUT1:Object at row " << row_number << " does not contain a 'name'! No dat file was generated.\n";
				break;
			case 2:
				std::clog << sheets_v[sheet_nr].name << "(" << row_number << ")  : File saving warning FDATOUT2:Could not create file for writing for object " << filename << "!\n";
				break;
			case 3:
				std::clog << sheets_v[sheet_nr].name << "(" << row_number << ") : File writing warning FDATOUT3:An error happened when writting on file for object " << filename << "! File may be corrupt.\n";
				break;
		}

		// time to set this filename as the one to be checked next
		last_filename = filename;
	}
}

/**
 * @brief Write the dat file on disk
 *
 * Simply dumps all data in dat_stream inside a file replacing it
 * if it already exists.
 *
 * @param filename Name of the file without extension
 * @param dat_stream String containing the whole dat file
 * @param append Whether it should append to existing file, if not it overwrites
 *
 * @return 0 if no errors
 * @return 1 if no filename was provided
 * @return 2 if file could not be opened
 * @return 3 if writing failed
 */
const unsigned int XLSX::writeDat(const std::string& filename, const std::string& dat_stream, const bool append)
{
	// std::cout << filename << ", append: " << append << std::endl;

	if (filename.length() == 0) {
		return 1;
	}

	// open file replacing if it already exists
	std::ofstream dat_file(filename + ".dat", (append ? std::ios::app : std::ios::trunc));

	if (!dat_file.is_open()) {
		return 2;
	}

	dat_file << (append ? "---\n" : "") << dat_stream;

	dat_file.close();

	if (dat_file.fail()) {
		return 3;
	}
	return 0;
}
