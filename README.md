# Simutrans DAT Sheet

This small software takes a specially crafted XLSX file (Microsoft Open Office XML Format) and creates DAT files for Simutrans (Simutrans Object Description Format).

The idea of the project is to allow pakset maintainers to more easily balance their paksets. Sheet editor software, like Excel and Calc, have functions, to automatically generate parameters for the game objects, and graphics, to visualise the data more clearly.

Compiled versions for Linux and Windows are available in the [Releases page](https://github.com/An-dz/datSheet/releases).

## Dependencies

* [pugixml](https://pugixml.org/)
* [libzip](https://libzip.org/)
* [libzippp] (libzip++ does no longer exist it seems)
* [ICU](http://site.icu-project.org/) (importing)
* [Windows SDK](https://developer.microsoft.com/windows/downloads/sdk-archive) (Windows)

To compile with MSVC, you just need to enable the VCPKG manifest and download the pugixml source files to an folder (and maybe change the include path).