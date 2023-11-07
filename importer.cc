#include <iostream>  // cout, cerr, clog
#include <sstream>   // ostringstream
#include <fstream>   // ifstream
#include <string>    // to_string
#include <cstring>   // strncmp, strlen, strrchr
#include <ctime>     // time, gmtime, strftime
#include <algorithm> // lower_bound, replace, transform
#include <unicode/ucsdet.h>
#include <unicode/ucnv.h>

#ifdef _WIN32
#include <windows.h>
#else
#include <sys/types.h>
#include <dirent.h>
#endif

#include "importer.hh"

/**
 * @brief Open an xlsx file
 *
 * An xlsx file is a normal zip file with multiple xmls inside.
 * This will open the zip for reding the files inside.
 *
 * @param filename Name of the spreadsheet file
 */
Importer::Importer(const std::string& filename)
{
	try {
		this->sheet = new libzippp::ZipArchive(filename);
		// open as writeable and replace everything
		this->sheet->open(libzippp::ZipArchive::Write);
		this->sheet->addEntry("_rels");
		this->sheet->addEntry("docProps");
		this->sheet->addEntry("xl");
		this->sheet->addEntry("xl/_rels");
		this->sheet->addEntry("xl/worksheets");
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
Importer::~Importer()
{
	sheet->close();
	delete sheet;
}

/**
 * @brief Add xml declaration
 *
 * Adds the first line of the xml file with its declaration.
 * It's always added for every single file.
 *
 * @param doc pugi::xml_document to add declaration
 */
void Importer::addXMLdeclaration(pugi::xml_document& doc)
{
	pugi::xml_node node = doc.append_child(pugi::node_declaration);
	pugi::xml_attribute attr = node.append_attribute("version");
	attr.set_value("1.0");
	attr = node.append_attribute("encoding");
	attr.set_value("UTF-8");
	attr = node.append_attribute("standalone");
	attr.set_value("yes");
}

/**
 * @brief Convert string into UTF-8
 *
 * Will attempt to convert an input string in whatever encoding
 * into UTF-8 encoding and will place the result into the passed
 * string stream.
 *
 * @param input  String in whatever encoding to be converted
 * @param output String stream where UTF-8 encoded result will be put
 *
 * @return bool telling if conversion was sucessful
 */
bool Importer::convertToUTF8(const std::string& input, std::stringstream& output)
{
	UErrorCode enc_status = U_ZERO_ERROR;
	UCharsetDetector *enc_detect = ucsdet_open(&enc_status);
	if (U_FAILURE(enc_status)) { return false; }

	ucsdet_setText(enc_detect, input.c_str(), input.length(), &enc_status);
	if (U_FAILURE(enc_status)) { ucsdet_close(enc_detect); return false; }

	// guess encoding
	const UCharsetMatch *enc_match = ucsdet_detect(enc_detect, &enc_status);
	if (U_FAILURE(enc_status)) { ucsdet_close(enc_detect); return false; }

	const char* enc_name = ucsdet_getName(enc_match, &enc_status);
	if (U_FAILURE(enc_status)) { ucsdet_close(enc_detect); return false; }

	UConverter *enc_cnv = ucnv_open(enc_name, &enc_status);
	if (U_FAILURE(enc_status)) { ucsdet_close(enc_detect); return false; }

	// Simutrans dat files are mostly ASCII, only comments can have special chars
	// so twice the size is probably more than enough
	size_t max_size = input.length() * 2;
	char *converted = new char(max_size);
//	char converted[max_size];

	// convert to UTF-8
	max_size = ucnv_toAlgorithmic(UCNV_UTF8, enc_cnv, converted, max_size, input.c_str(), input.length(), &enc_status);
	if (U_FAILURE(enc_status)) { ucsdet_close(enc_detect); return false; }

	// put result into stream
	output << std::string(converted, max_size) << '\0';

	ucsdet_close(enc_detect);
	return true;
}

/**
 * @brief Create an worksheet file
 *
 * Writes all dats passed in a worksheet file and saves it
 * in the xlsx file.
 *
 * @param dats Vector containing the location of all dats to write
 * @param dir Directory where the dat files are
 * @param index The index of the sheet, used for saving the correct file
 */
void Importer::createSheet(const std::vector<std::string>& dats, const std::string& dir, const unsigned int index)
{
	// start creating XML
	pugi::xml_document doc;
	addXMLdeclaration(doc);

	// worksheet main element
	pugi::xml_node worksheet = doc.append_child("worksheet");
	pugi::xml_attribute attr = worksheet.append_attribute("xmlns");
	attr.set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	attr = worksheet.append_attribute("xmlns:r");
	attr.set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	attr = worksheet.append_attribute("xmlns:mc");
	attr.set_value("http://schemas.openxmlformats.org/markup-compatibility/2006");
	attr = worksheet.append_attribute("mc:Ignorable");
	attr.set_value("x14ac");
	attr = worksheet.append_attribute("xmlns:x14ac");
	attr.set_value("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

	// views, freeze first column and row
	pugi::xml_node child1 = worksheet.append_child("sheetViews");
	child1 = child1.append_child("sheetView");
	// only one sheet is "open" and that's the first one
	// aka the sheet that shows when you open the xlsx file
	if (index == 1) {
		attr = child1.append_attribute("tabSelected");
		attr.set_value("1");
	}
	attr = child1.append_attribute("workbookViewId");
	attr.set_value("0");
	pugi::xml_node child2 = child1.append_child("pane");
	attr = child2.append_attribute("xSplit");
	attr.set_value("1");
	attr = child2.append_attribute("ySplit");
	attr.set_value("1");
	attr = child2.append_attribute("topLeftCell");
	attr.set_value("B2");
	attr = child2.append_attribute("activePane");
	attr.set_value("bottomRight");
	attr = child2.append_attribute("state");
	attr.set_value("frozen");
	child2 = child1.append_child("selection");
	attr = child2.append_attribute("pane");
	attr.set_value("topRight");
	attr = child2.append_attribute("activeCell");
	attr.set_value("B1");
	attr = child2.append_attribute("sqref");
	attr.set_value("B1");
	child2 = child1.append_child("selection");
	attr = child2.append_attribute("pane");
	attr.set_value("bottomLeft");
	attr = child2.append_attribute("activeCell");
	attr.set_value("A2");
	attr = child2.append_attribute("sqref");
	attr.set_value("A2");
	child2 = child1.append_child("selection");
	attr = child2.append_attribute("pane");
	attr.set_value("bottomRight");

	// set default row height
	child1 = worksheet.append_child("sheetFormatPr");
	attr = child1.append_attribute("defaultRowHeight");
	attr.set_value("15");

	// finally start adding what's in the dats
	child1 = worksheet.append_child("sheetData");
	pugi::xml_node child3;

	// parameter names to build row 1
	std::vector<std::string> parameters;
	// always start with name parameter
	parameters.push_back("name");
	parameters.push_back("filename");

	// skip first (parameter names) row as we don't know them yet
	// row 0 does not exist
	int row = 1;
	for (auto const& dat_name : dats) {
		// open file, read it all and put in string
		std::ifstream dat_file_open(dir + dat_name);
		std::string dat_buf;
		std::getline(dat_file_open, dat_buf, '\0');
		std::stringstream dat_file;

		// put converted string into stream
		if (convertToUTF8(dat_buf, dat_file)) {
			// list of column xml nodes
			std::vector<unsigned short> cols_list;
			pugi::xml_node nodes_list[702];
			std::string param;
			bool createRow = true;

			// add columns/paramater values
			while (std::getline(dat_file, param)) {
				// check for key = value
				std::stringstream param_stream(param);
				std::getline(param_stream, param, '=');

				// remove leading and trailing whitespaces
				std::string::size_type start = param.find_first_not_of(" \t");
				std::string::size_type end = param.find_last_not_of(" \t\n\r");
				// no checks needed, start always returns at least \r|\n
				param = param.substr((start != std::string::npos ? start : param.length()), (end == std::string::npos ? end : end + 1 - start));

				// object separation line, move to next row
				if (param.front() == '-') {
					// move to next row/object
					createRow = true;
					// restart columns/parameters list
					cols_list.clear();
				}
				// skip empty lines
				else if (param.size() > 1 && param.front() != '\r' && param.front() != '#') {
					std::string value;
					std::string type("s");

					// comment line
					if (param.front() == '#') {
						value = param.substr(1);
						param = "#";
					}
					else {
						// convert parameter to lowercase to merge things like Name & name
						std::transform(param.begin(), param.end(), param.begin(), ::tolower);
						std::getline(param_stream, value);

						if (value.length() > 0) {
							// remove leading and trailing whitespaces
							start = value.find_first_not_of(" \t");
							end = value.find_last_not_of(" \t\n\r");
							// no checks needed, start always returns at least \r|\n
							value = value.substr(start, (end == std::string::npos ? end : end + 1 - start));

							// number type
							if (value.find_first_not_of("0123456789") == std::string::npos) {
								type = "n";
							}
						}
					}

					if (value.length() > 0) {
						// string are saved in sharedStrings file
						if (type == "s") {
							value = std::to_string(findInVectorOrAdd(sharedStrings, value));
						}

						// create row
						if (createRow) {
							child2 = child1.append_child("row");
							attr = child2.append_attribute("r");
							attr.set_value(++row);
							createRow = false;
						}

						// get the column this value must be defined
						unsigned short col = findInVectorOrAdd(parameters, param);

						// get column letter code
						unsigned char div = col / 26;
						std::ostringstream cell_pos;
						if (div > 0) {
							// @ is 'A-1', that's because when div = 1 it has to add 'A'
							cell_pos << (unsigned char)(div + '@');
						}
						cell_pos << (unsigned char)((col % 26) + 'A') << row;

						// add the column
						auto col_lower = std::lower_bound(cols_list.begin(), cols_list.end(), col);

						// if column already exists we replace and alert
						if (col_lower != cols_list.end() && *col_lower == col) {
							pugi::xml_text value_node = nodes_list[col].child("v").text();
							std::clog << dir << dat_name << " : Value overwriten warning OV0:Parameter '" << param << "' overwritten." << std::endl;
							value_node.set(value.c_str());
						}
						else {
							if (col_lower == cols_list.end()) {
								nodes_list[col] = child2.append_child("c");
							}
							else {
								nodes_list[col] = child2.insert_child_before("c", nodes_list[*col_lower]);
							}

							cols_list.insert(col_lower, col);

							attr = nodes_list[col].append_attribute("r");
							attr.set_value(cell_pos.str().c_str());
							attr = nodes_list[col].append_attribute("t");
							attr.set_value(type.c_str());
							child3 = nodes_list[col].append_child("v");
							child3 = child3.append_child(pugi::node_pcdata);
							child3.set_value(value.c_str());
						}
					}
					else {
						std::clog << dir << dat_name << " : Value is null warning NV0:The following line seems to be invalid and was ignored:\n\t" << param << std::endl;
					}
				}
			}
		}
		else {
			std::clog << dir << dat_name << " : Encoding warning UE0:An error occurred while trying to detect file encoding. File was skipped. Saving it under a Unicode encoding will most likely fix this.";
		}
	}

	// prepend row 1
	child2 = child1.prepend_child("row");
	attr = child2.append_attribute("r");
	attr.set_value("1");
	unsigned short col = 0;
	// populate row 1 with parameter names
	for (auto const& param : parameters) {
		// get column letter code
		unsigned char div = col / 26;
		std::ostringstream cell_pos;
		if (div > 0) {
			cell_pos << (unsigned char)(div + '@');
		}
		cell_pos << (unsigned char)((col % 26) + 'A') << "1";

		std::string value = std::to_string(findInVectorOrAdd(sharedStrings, param));

		child3 = child2.append_child("c");
		attr = child3.append_attribute("r");
		attr.set_value(cell_pos.str().c_str());
		attr = child3.append_attribute("t");
		attr.set_value("s");
		child3 = child3.append_child("v");
		child3 = child3.append_child(pugi::node_pcdata);
		child3.set_value(value.c_str());

		col++;
	}

	std::string sheet_name("xl/worksheets/sheet" + std::to_string(index) + ".xml");
	std::ostringstream buffer;
	doc.save(buffer, "", pugi::format_raw);
//	sheet->add(libzip::source_buffer(buffer.str()), sheet_name, ZIP_FL_ENC_UTF_8);
	sheet->addData(sheet_name, buffer.str().c_str(), buffer.str().size());
}

/**
 * @brief Check if value is in vector and return index
 *
 * Will iterate over the vector to see if `value` is in the vector.
 * If `value` is not in the vector it's added.
 *
 * @param str_v Vector to be searched on
 * @param value Value to be searched for
 *
 * @return index where value is located
 */
unsigned int Importer::findInVectorOrAdd(std::vector<std::string>& str_v, const std::string& value)
{
	const unsigned int size = str_v.size();

	for (unsigned int i = 0; i < size; ++i) {
		if (str_v[i] == value) {
			return i;
		}
	}

	// if it was not found we add to it
	str_v.push_back(value);
	// index starts at 0, size is last_index + 1
	// thus we just need to return the old size
	return size;
}

/**
 * @brief Read dir and get dats and subfolders
 *
 * Will obtain all subfolders and dat files and place them
 * in vectors to then be used to create the sheets.
 *
 * @warn The function is recursive, it calls itself for each
 * sub-folder and so on, index is also updated automatically.
 *
 * @param dir_name Directory to analyse and create sheets
 * @param index Current sheet index
 */
void Importer::readDir(const std::string& dir_name, unsigned int& index)
{
	/* we get the list of subdirs and dat files now */
	/* so we can deal with each later */
	std::vector<std::string> dirs;
	std::vector<std::string> dats;

#ifdef _WIN32
	// Windows only
	HANDLE dir;
	WIN32_FIND_DATAA ent;
	char find_term[2048];
	std::snprintf(find_term, 2048, "%s/*", dir_name.c_str());

	// try starting it up and fail if no handle found
	if ((dir = FindFirstFileA(find_term, &ent)) == INVALID_HANDLE_VALUE) {
		std::ostringstream err_msg;
		err_msg << "WRD" << errno << ":" << strerror(errno);
		throw std::runtime_error(err_msg.str());
	}

	// do first to include result from FindFirstFile
	do {
		// skip those two
		if (std::strcmp(ent.cFileName, ".") && std::strcmp(ent.cFileName, "..")) {
			const char* extension = std::strrchr(ent.cFileName, '.');

			// directories are added to dirs list
			if (ent.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
				dirs.push_back(std::string(ent.cFileName, std::strlen(ent.cFileName)));
			}
			// files that have extension .dat are added to dats list
			else if (extension != NULL && !std::strcmp(extension, ".dat")) {
				dats.push_back(std::string(ent.cFileName, std::strlen(ent.cFileName)));
			}
		}
	} while (FindNextFileA(dir, &ent));

	FindClose(dir);
#else
	// Other platforms (Linux/OpenBSD)
	DIR *dir;
	struct dirent *ent;

	// failed opening directory
	if ((dir = opendir(dir_name.c_str())) == NULL) {
		std::ostringstream err_msg;
		err_msg << "URD" << errno << ":" << strerror(errno);
		throw std::runtime_error(err_msg.str());
	}

	// read directory
	while ((ent = readdir(dir)) != NULL) {
		// skip those two
		if (std::strcmp(ent->d_name, ".") && std::strcmp(ent->d_name, "..")) {
			const char* extension = std::strrchr(ent->d_name, '.');

			// directories are added to dirs list
			if (ent->d_type == DT_DIR) {
				dirs.push_back(std::string(ent->d_name, std::strlen(ent->d_name)));
			}
			// regular files that have extension .dat are added to dats list
			else if (ent->d_type == DT_REG && extension != NULL && !std::strcmp(extension, ".dat")) {
				dats.push_back(std::string(ent->d_name, std::strlen(ent->d_name)));
			}
		}
	}
	closedir(dir);
#endif

	// does current dir has an ending slash?
	// if it does not we will have to add it
	bool has_slash = (dir_name.find_last_of("\\/") == dir_name.size() - 1);
	bool is_root = dir_name.length() == root_dir_size;

	// create sheet file
	if (dats.size() > 0) {
		createSheet(dats, dir_name + (has_slash ? "" : "/"),  index++);
		if (!is_root) {
			std::string sheet_name = dir_name.substr(root_dir_size);
			std::replace(sheet_name.begin(), sheet_name.end(), '/',  ';');
			worksheets.push_back(sheet_name);
		}
		else {
			worksheets.push_back(";");
		}
	}

	// take into account the ending slash that will be added if not present
	if (is_root && !has_slash) {
		root_dir_size = root_dir_size + 1;
	}

	// enter sub-folders
	if (dirs.size() > 0) {
		for (auto const& dir : dirs) {
			readDir(dir_name + (has_slash ? "" : "/") + dir, index);
		}
	}
}

/**
 * @brief Starts the importing
 *
 * Imports a pakset structure into an xlsx file
 *
 * @param root_dir Root directory of the pakset
 */
void Importer::import(const std::string& root_dir)
{
	// used later for relative paths
	root_dir_size = root_dir.length();

	/*
	 * /xl/worksheets/sheet($index).xml
	 *
	 * Sheet files, each on its own xml file
	 */
	unsigned int index = 1;
	readDir(root_dir, index);

	/*
	 * /_rels/.rels
	 *
	 * Main relationships file, fixed data
	 * Defines where the properties and the workbook are
	 */
	std::ostringstream buffer("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
//	sheet->add(libzip::source_buffer(buffer.str()), "_rels/.rels", ZIP_FL_ENC_UTF_8);
	sheet->addData("_rels/.rels", buffer.str().c_str(), buffer.str().size());

	buffer.str("");
	pugi::xml_node node1, node2, node3;
	pugi::xml_attribute attr;

	/*
	 * /xl/sharedStrings.xml
	 *
	 * Holds all the strings, this is used to save space
	 * as multiple duplicate strings in the same file can
	 * point to the same string stored here.
	 */
	pugi::xml_document shared;
	addXMLdeclaration(shared);
	node1 = shared.append_child("sst");
	attr = node1.append_attribute("xmlns");
	attr.set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	for (auto const& string : sharedStrings) {
		node2 = node1.append_child("si");
		node2 = node2.append_child("t");
		node2 = node2.append_child(pugi::node_pcdata);
		node2.set_value(string.c_str());
	}

	shared.save(buffer, "", pugi::format_raw);
	//sheet->add(libzip::source_buffer(buffer.str()), "xl/sharedStrings.xml", ZIP_FL_ENC_UTF_8);
	sheet->addData("xl/sharedStrings.xml", buffer.str().c_str(), buffer.str().size());
	buffer.str("");

	/*
	 * /xl/_rels/workbook.xml.rels
	 *
	 * Workbook relationships file
	 * Defines where the workbook files are
	 */
	pugi::xml_document workbook_rels;
	addXMLdeclaration(workbook_rels);
	node1 = workbook_rels.append_child("Relationships");
	attr = node1.append_attribute("xmlns");
	attr.set_value("http://schemas.openxmlformats.org/package/2006/relationships");
	// id must continue for sharedStrings
	unsigned int id = 0;
	while (id < worksheets.size()) {
		id++;
		node2 = node1.append_child("Relationship");
		attr = node2.append_attribute("Id");
		attr.set_value(std::string("rId" + std::to_string(id)).c_str());
		attr = node2.append_attribute("Type");
		attr.set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
		attr = node2.append_attribute("Target");
		attr.set_value(std::string("worksheets/sheet" + std::to_string(id) + ".xml").c_str());
	}
	node2 = node1.append_child("Relationship");
	attr = node2.append_attribute("Id");
	attr.set_value(std::string("rId" + std::to_string(++id)).c_str());
	attr = node2.append_attribute("Type");
	attr.set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
	attr = node2.append_attribute("Target");
	attr.set_value("sharedStrings.xml");

	workbook_rels.save(buffer, "", pugi::format_raw);
//	sheet->add(libzip::source_buffer(buffer.str()), "xl/_rels/workbook.xml.rels", ZIP_FL_ENC_UTF_8);
	sheet->addData("xl/_rels/workbook.xml.rels", buffer.str().c_str(), buffer.str().size());
	buffer.str("");

	/*
	 * /[Content_Types].xml
	 *
	 * Defines the types of all files in the xlsx
	 */
	pugi::xml_document types;
	addXMLdeclaration(types);
	node1 = types.append_child("Types");
	attr = node1.append_attribute("xmlns");
	attr.set_value("http://schemas.openxmlformats.org/package/2006/content-types");
	// rels files
	node2 = node1.append_child("Default");
	attr = node2.append_attribute("Extension");
	attr.set_value("rels");
	attr = node2.append_attribute("ContentType");
	attr.set_value("application/vnd.openxmlformats-package.relationships+xml");
	// xml files
	node2 = node1.append_child("Default");
	attr = node2.append_attribute("Extension");
	attr.set_value("xml");
	attr = node2.append_attribute("ContentType");
	attr.set_value("application/xml");
	// workbook file
	node2 = node1.append_child("Override");
	attr = node2.append_attribute("PartName");
	attr.set_value("/xl/workbook.xml");
	attr = node2.append_attribute("ContentType");
	attr.set_value("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
	// strings file
	node2 = node1.append_child("Override");
	attr = node2.append_attribute("PartName");
	attr.set_value("/xl/sharedStrings.xml");
	attr = node2.append_attribute("ContentType");
	attr.set_value("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
	// core properties file
	node2 = node1.append_child("Override");
	attr = node2.append_attribute("PartName");
	attr.set_value("/docProps/core.xml");
	attr = node2.append_attribute("ContentType");
	attr.set_value("application/vnd.openxmlformats-package.core-properties+xml");
	// app properties file
	node2 = node1.append_child("Override");
	attr = node2.append_attribute("PartName");
	attr.set_value("/docProps/app.xml");
	attr = node2.append_attribute("ContentType");
	attr.set_value("application/vnd.openxmlformats-officedocument.extended-properties+xml");
	// sheet files
	for (unsigned int id = 0; id < worksheets.size(); ) {
		node2 = node1.append_child("Override");
		attr = node2.append_attribute("PartName");
		attr.set_value(std::string("/xl/worksheets/sheet" + std::to_string(++id) + ".xml").c_str());
		attr = node2.append_attribute("ContentType");
		attr.set_value("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
	}

	types.save(buffer, "", pugi::format_raw);
//	sheet->add(libzip::source_buffer(buffer.str()), "[Content_Types].xml", ZIP_FL_ENC_UTF_8);
	sheet->addData("[Content_Types].xml", buffer.str().c_str(), buffer.str().size());
	buffer.str("");

	/*
	 * /xl/workbook.xml
	 *
	 * Defines the workbook, contains the names of the sheets
	 */
	pugi::xml_document workbook;
	addXMLdeclaration(workbook);
	node1 = workbook.append_child("workbook");
	attr = node1.append_attribute("xmlns");
	attr.set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	attr = node1.append_attribute("xmlns:r");
	attr.set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	node2 = node1.append_child("sheets");
	id = 1;
	for (auto const& sheet_name : worksheets) {
		node3 = node2.append_child("sheet");
		attr = node3.append_attribute("name");
		attr.set_value(sheet_name.c_str());
		attr = node3.append_attribute("sheetId");
		attr.set_value(id);
		attr = node3.append_attribute("r:id");
		attr.set_value(std::string("rId" + std::to_string(id)).c_str());
		id++;
	}

	workbook.save(buffer, "", pugi::format_raw);
//	sheet->add(libzip::source_buffer(buffer.str()), "xl/workbook.xml", ZIP_FL_ENC_UTF_8);
	sheet->addData("xl/workbook.xml", buffer.str().c_str(), buffer.str().size());
	buffer.str("");

	/*
	 * /docProps/app.xml
	 *
	 * Properties about the application that created the xlsx
	 * (aka ourselves) and the xlsx itself
	 */
	pugi::xml_document app;
	addXMLdeclaration(app);
	node1 = app.append_child("Properties");
	attr = node1.append_attribute("xmlns");
	attr.set_value("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
	attr = node1.append_attribute("xmlns:vt");
	attr.set_value("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
	// that's us :)
	node2 = node1.append_child("Application");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value("datSheet");
	// no passwords
	node2 = node1.append_child("DocSecurity");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value("0");
	// preview thumbnail, true=resize, false=crop
	node2 = node1.append_child("ScaleCrop");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value("false");
	// how the sheets are organised in the document
	node2 = node1.append_child("HeadingPairs");
	node2 = node2.append_child("vt:vector");
	attr = node2.append_child("size");
	attr.set_value("2");
	attr = node2.append_child("baseType");
	attr.set_value("variant");
	node3 = node2.append_child("vt:variant");
	node3 = node3.append_child("vt:lpstr");
	node3 = node3.append_child(pugi::node_pcdata);
	node3.set_value("Sheets"); // just a name for the group
	node3 = node2.append_child("vt:variant");
	node3 = node3.append_child("vt:i4");
	node3 = node3.append_child(pugi::node_pcdata);
	const char* size = std::to_string(worksheets.size()).c_str();
	node3.set_value(size); // number of sheets
	// names of the sheets
	node2 = node1.append_child("TitlesOfParts");
	node2 = node2.append_child("vt:vector");
	attr = node2.append_attribute("size");
	attr.set_value(size);
	attr = node2.append_attribute("baseType");
	attr.set_value("lpstr");
	for (auto const& sheet_name : worksheets) {
		node3 = node2.append_child("vt:lpstr");
		node3 = node3.append_child(pugi::node_pcdata);
		node3.set_value(sheet_name.c_str());
	}
	// if links are up-to-date, we don't have them
	node2 = node1.append_child("LinksUpToDate");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value("false");
	// if this doc is being shared (multiple concurrent editors), obviously not
	node2 = node1.append_child("SharedDoc");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value("false");
	// if hyperlinks must be updated by next app
	node2 = node1.append_child("HyperlinksChanged");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value("false");
	// our version
	node2 = node1.append_child("AppVersion");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value(VERSION);

	app.save(buffer, "", pugi::format_raw);
//	sheet->add(libzip::source_buffer(buffer.str()), "docProps/app.xml", ZIP_FL_ENC_UTF_8);
	sheet->addData("docProps/app.xml", buffer.str().c_str(), buffer.str().size());
	buffer.str("");

	/*
	 * /docProps/core.xml
	 *
	 * Core properties, app and file type independent
	 */
	pugi::xml_document core;
	addXMLdeclaration(core);
	node1 = core.append_child("cp:coreProperties");
	attr = node1.append_attribute("xmlns:cp");
	attr.set_value("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
	attr = node1.append_attribute("xmlns:dc");
	attr.set_value("http://purl.org/dc/elements/1.1/");
	attr = node1.append_attribute("xmlns:dcterms");
	attr.set_value("http://purl.org/dc/terms/");
	attr = node1.append_attribute("xmlns:dcmitype");
	attr.set_value("http://purl.org/dc/dcmitype/");
	attr = node1.append_attribute("xmlns:xsi");
	attr.set_value("http://www.w3.org/2001/XMLSchema-instance");
	// name of the file
	node2 = node1.append_child("dc:title");
	node2 = node2.append_child(pugi::node_pcdata);

	std::string pakname = root_dir;
	if (root_dir.find_last_of("\\/") == root_dir_size - 1) {
		pakname = pakname.substr(0, root_dir_size - 1);
	}
	pakname = pakname.substr(pakname.find_last_of("\\/") + 1);
	node2.set_value(pakname.c_str());
	// name of the author
	node2 = node1.append_child("dc:creator");
	node2 = node2.append_child(pugi::node_pcdata);

	std::string team_name = pakname + " team";
	node2.set_value(team_name.c_str());
	node2 = node1.append_child("cp:lastModifiedBy");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value(team_name.c_str());

	// creation time
	char time[21];
	std::time_t now = std::time(nullptr);
	std::strftime(time, 21, "%Y-%m-%dT%H:%M:%SZ", std::gmtime(&now));

	node2 = node1.append_child("dcterms:created");
	attr = node2.append_attribute("xsi:type");
	attr.set_value("dcterms:W3CDTF");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value(time);
	node2 = node1.append_child("dcterms:modified");
	attr = node2.append_attribute("xsi:type");
	attr.set_value("dcterms:W3CDTF");
	node2 = node2.append_child(pugi::node_pcdata);
	node2.set_value(time);

	core.save(buffer, "", pugi::format_raw);
//	sheet->add(libzip::source_buffer(buffer.str()), "docProps/core.xml", ZIP_FL_ENC_UTF_8);
	sheet->addData("docProps/core.xml", buffer.str().c_str(), buffer.str().size());
}
