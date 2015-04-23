#include <Rcpp.h>
#include "XlsxWorkBook.h"
using namespace Rcpp;

// [[Rcpp::export]]
CharacterVector xlsx_sheets(std::string path) {
  return XlsxWorkBook(path).sheets();
}

// [[Rcpp::export]]
CharacterVector xlsx_namedranges(std::string path) {
  return XlsxWorkBook(path).namedranges();
}

// [[Rcpp::export]]
std::vector<std::string> xlsx_strings(std::string path) {
  return XlsxWorkBook(path).stringTable();
}

// [[Rcpp::export]]
std::set<int> xlsx_date_styles(std::string path) {
  return XlsxWorkBook(path).dateStyles();
}
