///////////////////////////////////////////////////////////////////////////////
// NAME:            main.c
//
// AUTHOR:          Ethan D. Twardy <ethan.twardy@gmail.com>
//
// DESCRIPTION:     Entrypoint for the program
//
// CREATED:         10/28/2021
//
// LAST EDITED:     10/28/2021
//
// Copyright 2021, Ethan D. Twardy
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
////

#include <stdbool.h>
#include <stdio.h>
#include <stdlib.h>

#define R_NO_REMAP
#include <R.h>
#include <Rinternals.h>
#include <Rembedded.h>

#include "config.h"
#include "Budget.h"

typedef enum { INTEREST, DISCOUNT, FORCE } Type;

typedef struct { bool ok; double value; } DoubleResult;
typedef struct { bool ok; const char* value; } StringResult;

StringResult type_to_string(Type type) {
  StringResult result = {0};

  static const char* s_interest = "interest";
  static const char* s_discount = "discount";
  static const char* s_force = "force";

  switch (type) {
  case INTEREST: result.ok = true, result.value = s_interest;
  case DISCOUNT: result.ok = true, result.value = s_discount;
  case FORCE: result.ok = true, result.value = s_force;
  default: break;
  }

  return result;
}

const char* bool_to_string(bool value) {
  static const char* s_true = "true";
  static const char* s_false = "false";

  if (value) {
    return s_true;
  }

  return s_false;
}

void library(const char *name)
{
  SEXP e;

  PROTECT(e = Rf_lang2(Rf_install("library"), Rf_mkString(name)));
  R_tryEval(e, R_GlobalEnv, NULL);
  UNPROTECT(1);
}

// Invoke the rate.conv method from the 'FinancialMath' package.
DoubleResult rate_conv(double rate, int convertibleTimesPerYear, Type type,
  int numberOfTimesConvertible)
{
  DoubleResult result = {0};
  StringResult type_string = type_to_string(type);
  if (!type_string.ok) {
    return result;
  }

  SEXP function;
  // https://github.com/hadley/r-internals/blob/master/pairlists.md
  PROTECT(function = Rf_lang5(Rf_install("rate.conv"),
      Rf_ScalarReal(rate), Rf_ScalarInteger(convertibleTimesPerYear),
      Rf_mkString(type_string.value),
      Rf_ScalarInteger(numberOfTimesConvertible)));

  int errorOccurred;
  SEXP evalResult = R_tryEval(function, R_GlobalEnv, &errorOccurred);
  UNPROTECT(1);

  result.ok = !!errorOccurred;
  result.value = *REAL(evalResult);
  return result;
}

int main() {
  setenv("R_HOME", ENV_R_HOME, 1);

  const int r_argc = 2;
  char* r_argv[] = { "R", "--silent" };
  Rf_initEmbeddedR(r_argc, r_argv);

  library("FinancialMath");
  DoubleResult result = rate_conv(0.0675, 12, INTEREST, 1);
  printf("Result: { %s, %f }\n", bool_to_string(result.ok), result.value);

  Rf_endEmbeddedR(0);
}

///////////////////////////////////////////////////////////////////////////////
