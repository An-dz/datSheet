#include <iostream>     // cout, cerr, clog, left, endl
#include <iomanip>      // setw
#include <cstring>      // strncmp
#include "xlsx.hh"      // XLSX parser
#include "importer.hh"  // XLSX importer

int main(int argc, char const *argv[])
{
	int option = 0;
	int files[256];
	int num_files = 0;

	// check passed arguments
	for (int i = 1; i < argc; ++i) {
		if (!std::strncmp(argv[i], "-V", 3) || !std::strncmp(argv[i], "--version", 10)) {
			option |= 1;
		}
		else if (!std::strncmp(argv[i], "-h", 3) || !std::strncmp(argv[i], "--help", 7)) {
			option |= 2;
		}
		else if (!std::strncmp(argv[i], "-i", 3) || !std::strncmp(argv[i], "--import", 9)) {
			option |= 4;
		}
		else if (argv[i][0] != '-') {
			files[num_files++] = i;
		}
	}

	// if no arguments were passed and no option was chosen
	if ((num_files == 0 && option == 0) || (option == 4 && num_files < 2)) {
		std::clog << "datSheet : No file error NFN:No file specified!\n";
		return EXIT_FAILURE;
	}

	// if --help was seleced
	if (option > 1 && option != 4) {
		std::cout << "usage:  datSheet [dir] <file(s)>\n\noptions:\n   " << std::left << std::setw(15) << "-i --import" << "Create sheet file from one directory\n" << std::setw(15) << "-h --help" << "Display this help text\n   " << std::setw(15) << "-V --version" << "Print version\n\nsupported file types: XLSX\n\nproject homepage: <https://github.com/An-dz/datSheet>\n";
		return EXIT_SUCCESS;
	}
	// if --version was selected
	if (option > 0 && option != 4) {
		std::cout << "Simutrans datSheet " << VERSION << "\n   Copyright (c) 2018 Andre' Zanghelini (An_dz)\n   Project homepage: <https://github.com/An-dz/datSheet>\n\nA big thanks to the following libraries:\npugixml <https://pugixml.org>\n   Copyright (c) 2006-2018 Arseny Kapoulkine.\nlibzip <https://libzip.org/>\n   Copyright (c) 1999-2018 Dieter Baron and Thomas Klausner\nlibzip++ <http://hg.markand.fr/libzip>\n   Copyright (c) 2013-2018 David Demelier <markand@malikania.fr>\nICU <http://site.icu-project.org/>\n   Copyright (c) 1991-2018 Unicode, Inc. All rights reserved.\n";
		return EXIT_SUCCESS;
	}

	try {
		if (option == 0) {
			for (int i = 0; i < num_files; ++i) {
				XLSX xlsx(argv[files[i]]);
				xlsx.parse();
			}
		}
		else {
			Importer xlsx(argv[files[1]]);
			xlsx.import(argv[files[0]]);
		}
		std::cout << "Finished without errors.\n";
	} catch (const std::runtime_error& e) {
		std::cerr << "datSheet : error " << e.what() << std::endl;
		return EXIT_FAILURE;
	}
}
