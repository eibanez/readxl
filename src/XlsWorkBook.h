#ifndef READXL_XLSWORKBOOK_
#define READXL_XLSWORKBOOK_

#include <Rcpp.h>
#include <libxls/xls.h>
#include "CellType.h"

inline std::string normalizePath(std::string path) {
  Rcpp::Environment baseEnv = Rcpp::Environment::base_env();
  Rcpp::Function normalizePath = baseEnv["normalizePath"];
  return Rcpp::as<std::string>(normalizePath(path, "/", true));
}

class XlsWorkSheet;

typedef std::map<int,std::string> FormatMap;

class XlsWorkBook {
  std::string path_;
  xls::xlsWorkBook* pWB_;

public:

  XlsWorkBook(std::string path) {
    path_ = normalizePath(path);
    pWB_ = xls::xls_open(path_.c_str(), "UTF-8");
    if (pWB_ == NULL)
      Rcpp::stop("Failed to open %s", path);
  }

  ~XlsWorkBook() {
    try {
      xls_close_WB(pWB_);
    } catch(...) {}
  }

  xls::xlsWorkBook* workbook() const {
    return pWB_;
  }

  int nSheets() const {
    return pWB_->sheets.count;
  }

  int nDnames() const {
    return pWB_->dnames.count;
  }

  Rcpp::CharacterVector sheets() const {
    Rcpp::CharacterVector sheets(nSheets());

    for (int i = 0; i < nSheets(); ++i) {
      sheets[i] = Rf_mkCharCE((char*) pWB_->sheets.sheet[i].name, CE_UTF8);
    }

    return sheets;
  }

  Rcpp::DataFrame defined_names() const {
    Rcpp::CharacterVector names(nDnames());
    Rcpp::CharacterVector sheet(nDnames());
    Rcpp::NumericVector   row1(nDnames());
    Rcpp::NumericVector   row2(nDnames());
    Rcpp::NumericVector   col1(nDnames());
    Rcpp::NumericVector   col2(nDnames());

    for (int i = 0; i < nDnames(); ++i) {
      names[i] = Rf_mkCharCE((char*) pWB_->dnames.dname[i].name, CE_UTF8);
      sheet[i] = Rf_mkCharCE((char*) pWB_->sheets.sheet[pWB_->dnames.dname[i].sheet].name, CE_UTF8);
      row1[i]  = pWB_->dnames.dname[i].row1;
      row2[i]  = pWB_->dnames.dname[i].row2;
      col1[i]  = pWB_->dnames.dname[i].col1;
      col2[i]  = pWB_->dnames.dname[i].col2;
    }

    return Rcpp::DataFrame::create(Rcpp::Named("name")  = names,
                                   Rcpp::Named("sheet") = sheet,
                                   Rcpp::Named("row1")  = row1 + 1,
                                   Rcpp::Named("row2")  = row2 + 1,
                                   Rcpp::Named("col1")  = col1 + 1,
                                   Rcpp::Named("col2")  = col2 + 1);
  }


  int nFormats() const {
    return pWB_->formats.count;
  }

  FormatMap formats() const {
    std::map<int, std::string> formats;

    for (int i = 0; i < nFormats(); ++i) {
      xls::st_format::st_format_data format = pWB_->formats.format[i];
      std::string value((char*) pWB_->formats.format[i].value);

      formats.insert(std::make_pair(format.index, value));
    }

    return formats;
  }

  std::set<int> customDateFormats() const {
    std::set<int> dateFormats;

    for (int i = 0; i < nFormats(); ++i) {
      xls::st_format::st_format_data format = pWB_->formats.format[i];
      std::string value((char*) format.value);

      if (isDateFormat(value))
        dateFormats.insert(format.index);
    }

    return dateFormats;
  }


  XlsWorkSheet sheet(std::string name);
  XlsWorkSheet sheet(int i);

};

#endif
