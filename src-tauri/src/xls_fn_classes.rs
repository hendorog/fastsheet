//! Per-function parameter operand classes for BIFF formula
//! encoding, generated from Apache POI 5.5.1
//! `org/apache/poi/ss/formula/function/functionMetadata.txt`
//! (itself derived from the OpenOffice excelfileformat spec).
//! Index = BIFF iftab; value = operand-class byte per parameter
//! (0x20 = reference, 0x40 = value, 0x60 = array). For calls
//! with more args than listed classes, the last class repeats
//! (varargs convention). `None` / empty = zero-arg function or
//! absent from the table; the encoder falls back to value class
//! for its args.

pub(crate) static FN_PARAM_CLASSES: [Option<&[u8]>; 369] = [
    Some(&[0x20]), // 0 COUNT
    Some(&[0x40, 0x20, 0x20]), // 1 IF
    Some(&[0x40]), // 2 ISNA
    Some(&[0x40]), // 3 ISERROR
    Some(&[0x20]), // 4 SUM
    Some(&[0x20]), // 5 AVERAGE
    Some(&[0x20]), // 6 MIN
    Some(&[0x20]), // 7 MAX
    Some(&[0x20]), // 8 ROW
    Some(&[0x20]), // 9 COLUMN
    Some(&[]), // 10 NA
    Some(&[0x40, 0x20]), // 11 NPV
    Some(&[0x20]), // 12 STDEV
    Some(&[0x40, 0x40]), // 13 DOLLAR
    Some(&[0x40, 0x40, 0x40]), // 14 FIXED
    Some(&[0x40]), // 15 SIN
    Some(&[0x40]), // 16 COS
    Some(&[0x40]), // 17 TAN
    Some(&[0x40]), // 18 ATAN
    Some(&[]), // 19 PI
    Some(&[0x40]), // 20 SQRT
    Some(&[0x40]), // 21 EXP
    Some(&[0x40]), // 22 LN
    Some(&[0x40]), // 23 LOG10
    Some(&[0x40]), // 24 ABS
    Some(&[0x40]), // 25 INT
    Some(&[0x40]), // 26 SIGN
    Some(&[0x40, 0x40]), // 27 ROUND
    Some(&[0x40, 0x20, 0x20]), // 28 LOOKUP
    Some(&[0x20, 0x40, 0x40, 0x40]), // 29 INDEX
    Some(&[0x40, 0x40]), // 30 REPT
    Some(&[0x40, 0x40, 0x40]), // 31 MID
    Some(&[0x40]), // 32 LEN
    Some(&[0x40]), // 33 VALUE
    Some(&[]), // 34 TRUE
    Some(&[]), // 35 FALSE
    Some(&[0x20]), // 36 AND
    Some(&[0x20]), // 37 OR
    Some(&[0x40]), // 38 NOT
    Some(&[0x40, 0x40]), // 39 MOD
    Some(&[0x20, 0x20, 0x20]), // 40 DCOUNT
    Some(&[0x20, 0x20, 0x20]), // 41 DSUM
    Some(&[0x20, 0x20, 0x20]), // 42 DAVERAGE
    Some(&[0x20, 0x20, 0x20]), // 43 DMIN
    Some(&[0x20, 0x20, 0x20]), // 44 DMAX
    Some(&[0x20, 0x20, 0x20]), // 45 DSTDEV
    Some(&[0x20]), // 46 VAR
    Some(&[0x20, 0x20, 0x20]), // 47 DVAR
    Some(&[0x40, 0x40]), // 48 TEXT
    Some(&[0x20, 0x20, 0x40, 0x40]), // 49 LINEST
    Some(&[0x20, 0x20, 0x20, 0x40]), // 50 TREND
    Some(&[0x20, 0x20, 0x40, 0x40]), // 51 LOGEST
    Some(&[0x20, 0x20, 0x20, 0x40]), // 52 GROWTH
    Some(&[0x20]), // 53 GOTO
    None, // 54
    Some(&[0x40]), // 55 RETURN
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 56 PV
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 57 FV
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 58 NPER
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 59 PMT
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40, 0x40]), // 60 RATE
    Some(&[0x60, 0x40, 0x40]), // 61 MIRR
    Some(&[0x60, 0x40]), // 62 IRR
    Some(&[]), // 63 RAND
    Some(&[0x40, 0x20, 0x20]), // 64 MATCH
    Some(&[0x40, 0x40, 0x40]), // 65 DATE
    Some(&[0x40, 0x40, 0x40]), // 66 TIME
    Some(&[0x40]), // 67 DAY
    Some(&[0x40]), // 68 MONTH
    Some(&[0x40]), // 69 YEAR
    Some(&[0x40, 0x40]), // 70 WEEKDAY
    Some(&[0x40]), // 71 HOUR
    Some(&[0x40]), // 72 MINUTE
    Some(&[0x40]), // 73 SECOND
    Some(&[]), // 74 NOW
    Some(&[0x20]), // 75 AREAS
    Some(&[0x60]), // 76 ROWS
    Some(&[0x60]), // 77 COLUMNS
    Some(&[0x20, 0x40, 0x40, 0x40, 0x40]), // 78 OFFSET
    Some(&[0x40, 0x20]), // 79 ABSREF
    Some(&[0x40, 0x40]), // 80 RELREF
    Some(&[0x40, 0x40, 0x20]), // 81 ARGUMENT
    Some(&[0x40, 0x40, 0x40]), // 82 SEARCH
    Some(&[0x60]), // 83 TRANSPOSE
    Some(&[0x40, 0x20]), // 84 ERROR
    Some(&[]), // 85 STEP
    Some(&[0x40]), // 86 TYPE
    None, // 87
    None, // 88
    None, // 89
    None, // 90
    None, // 91
    None, // 92
    None, // 93
    None, // 94
    None, // 95
    None, // 96
    Some(&[0x40, 0x40]), // 97 ATAN2
    Some(&[0x40]), // 98 ASIN
    Some(&[0x40]), // 99 ACOS
    Some(&[0x40, 0x20]), // 100 CHOOSE
    Some(&[0x40, 0x20, 0x20, 0x40]), // 101 HLOOKUP
    Some(&[0x40, 0x20, 0x20, 0x40]), // 102 VLOOKUP
    None, // 103
    None, // 104
    Some(&[0x20]), // 105 ISREF
    None, // 106
    None, // 107
    None, // 108
    Some(&[0x40, 0x40]), // 109 LOG
    Some(&[0x40, 0x40, 0x40, 0x40]), // 110 EXEC
    Some(&[0x40]), // 111 CHAR
    Some(&[0x40]), // 112 LOWER
    Some(&[0x40]), // 113 UPPER
    Some(&[0x40]), // 114 PROPER
    Some(&[0x40, 0x40]), // 115 LEFT
    Some(&[0x40, 0x40]), // 116 RIGHT
    Some(&[0x40, 0x40]), // 117 EXACT
    Some(&[0x40]), // 118 TRIM
    Some(&[0x40, 0x40, 0x40, 0x40]), // 119 REPLACE
    Some(&[0x40, 0x40, 0x40, 0x40]), // 120 SUBSTITUTE
    Some(&[0x40]), // 121 CODE
    None, // 122
    None, // 123
    Some(&[0x40, 0x40, 0x40]), // 124 FIND
    Some(&[0x40, 0x20]), // 125 CELL
    Some(&[0x40]), // 126 ISERR
    Some(&[0x40]), // 127 ISTEXT
    Some(&[0x40]), // 128 ISNUMBER
    Some(&[0x40]), // 129 ISBLANK
    Some(&[0x20]), // 130 T
    Some(&[0x20]), // 131 N
    None, // 132
    None, // 133
    None, // 134
    None, // 135
    None, // 136
    None, // 137
    None, // 138
    None, // 139
    Some(&[0x40]), // 140 DATEVALUE
    Some(&[0x40]), // 141 TIMEVALUE
    Some(&[0x40, 0x40, 0x40]), // 142 SLN
    Some(&[0x40, 0x40, 0x40, 0x40]), // 143 SYD
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 144 DDB
    None, // 145
    None, // 146
    None, // 147
    Some(&[0x40, 0x40]), // 148 INDIRECT
    None, // 149
    Some(&[0x40, 0x20, 0x20]), // 150 CALL
    None, // 151
    None, // 152
    None, // 153
    None, // 154
    None, // 155
    None, // 156
    None, // 157
    None, // 158
    None, // 159
    None, // 160
    None, // 161
    Some(&[0x40]), // 162 CLEAN
    Some(&[0x60]), // 163 MDETERM
    Some(&[0x60]), // 164 MINVERSE
    Some(&[0x60, 0x60]), // 165 MMULT
    None, // 166
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40, 0x40]), // 167 IPMT
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40, 0x40]), // 168 PPMT
    Some(&[0x20]), // 169 COUNTA
    None, // 170
    None, // 171
    None, // 172
    None, // 173
    None, // 174
    None, // 175
    None, // 176
    None, // 177
    None, // 178
    None, // 179
    None, // 180
    None, // 181
    None, // 182
    Some(&[0x20]), // 183 PRODUCT
    Some(&[0x40]), // 184 FACT
    Some(&[0x40, 0x20]), // 185 GET.CELL
    Some(&[0x40]), // 186 GET.WORKSPACE
    Some(&[0x40, 0x40]), // 187 GET.WINDOW
    Some(&[0x40, 0x40]), // 188 GET.DOCUMENT
    Some(&[0x20, 0x20, 0x20]), // 189 DPRODUCT
    Some(&[0x40]), // 190 ISNONTEXT
    None, // 191
    None, // 192
    Some(&[0x20]), // 193 STDEVP
    Some(&[0x20]), // 194 VARP
    Some(&[0x20, 0x20, 0x20]), // 195 DSTDEVP
    Some(&[0x20, 0x20, 0x20]), // 196 DVARP
    Some(&[0x40, 0x40]), // 197 TRUNC
    Some(&[0x40]), // 198 ISLOGICAL
    Some(&[0x20, 0x20, 0x20]), // 199 DCOUNTA
    None, // 200
    None, // 201
    None, // 202
    None, // 203
    Some(&[0x40, 0x40]), // 204 USDOLLAR
    Some(&[0x40, 0x40, 0x40]), // 205 FINDB
    Some(&[0x40, 0x40, 0x40]), // 206 SEARCHB
    Some(&[0x40, 0x40, 0x40, 0x40]), // 207 REPLACEB
    Some(&[0x40, 0x40]), // 208 LEFTB
    Some(&[0x40, 0x40]), // 209 RIGHTB
    Some(&[0x40, 0x40, 0x40]), // 210 MIDB
    Some(&[0x40]), // 211 LENB
    Some(&[0x40, 0x40]), // 212 ROUNDUP
    Some(&[0x40, 0x40]), // 213 ROUNDDOWN
    Some(&[0x40]), // 214 ASC
    Some(&[0x40]), // 215 DBCS
    Some(&[0x40, 0x20, 0x40]), // 216 RANK
    None, // 217
    None, // 218
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 219 ADDRESS
    Some(&[0x40, 0x40, 0x40]), // 220 DAYS360
    Some(&[]), // 221 TODAY
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40, 0x40, 0x40]), // 222 VDB
    None, // 223
    None, // 224
    Some(&[]), // 225 END.IF
    None, // 226
    Some(&[0x20]), // 227 MEDIAN
    Some(&[0x60]), // 228 SUMPRODUCT
    Some(&[0x40]), // 229 SINH
    Some(&[0x40]), // 230 COSH
    Some(&[0x40]), // 231 TANH
    Some(&[0x40]), // 232 ASINH
    Some(&[0x40]), // 233 ACOSH
    Some(&[0x40]), // 234 ATANH
    Some(&[0x20, 0x20, 0x20]), // 235 DGET
    None, // 236
    None, // 237
    Some(&[]), // 238 LAST.ERROR
    None, // 239
    None, // 240
    None, // 241
    None, // 242
    None, // 243
    Some(&[0x40]), // 244 INFO
    None, // 245
    None, // 246
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 247 DB
    None, // 248
    None, // 249
    None, // 250
    None, // 251
    Some(&[0x20, 0x20]), // 252 FREQUENCY
    None, // 253
    None, // 254
    None, // 255
    None, // 256
    Some(&[0x40]), // 257 EVALUATE
    None, // 258
    None, // 259
    None, // 260
    Some(&[0x40]), // 261 ERROR.TYPE
    Some(&[0x40]), // 262 APP.TITLE
    Some(&[0x40]), // 263 WINDOW.TITLE
    Some(&[0x40, 0x40]), // 264 SAVE.TOOLBAR
    Some(&[0x40, 0x40, 0x40]), // 265 ENABLE.TOOL
    Some(&[0x40, 0x40, 0x40]), // 266 PRESS.TOOL
    Some(&[0x40, 0x40, 0x40]), // 267 REGISTER.ID
    Some(&[0x40, 0x40]), // 268 GET.WORKBOOK
    Some(&[0x20]), // 269 AVEDEV
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 270 BETADIST
    Some(&[0x40]), // 271 GAMMALN
    Some(&[0x40, 0x40, 0x40, 0x40, 0x40]), // 272 BETAINV
    Some(&[0x40, 0x40, 0x40, 0x40]), // 273 BINOMDIST
    Some(&[0x40, 0x40]), // 274 CHIDIST
    Some(&[0x40, 0x40]), // 275 CHIINV
    Some(&[0x40, 0x40]), // 276 COMBIN
    Some(&[0x40, 0x40, 0x40]), // 277 CONFIDENCE
    Some(&[0x40, 0x40, 0x40]), // 278 CRITBINOM
    Some(&[0x40]), // 279 EVEN
    Some(&[0x40, 0x40, 0x40]), // 280 EXPONDIST
    Some(&[0x40, 0x40, 0x40]), // 281 FDIST
    Some(&[0x40, 0x40, 0x40]), // 282 FINV
    Some(&[0x40]), // 283 FISHER
    Some(&[0x40]), // 284 FISHERINV
    Some(&[0x40, 0x40]), // 285 FLOOR
    Some(&[0x40, 0x40, 0x40, 0x40]), // 286 GAMMADIST
    Some(&[0x40, 0x40, 0x40]), // 287 GAMMAINV
    Some(&[0x40, 0x40]), // 288 CEILING
    Some(&[0x40, 0x40, 0x40, 0x40]), // 289 HYPGEOMDIST
    Some(&[0x40, 0x40, 0x40]), // 290 LOGNORMDIST
    Some(&[0x40, 0x40, 0x40]), // 291 LOGINV
    Some(&[0x40, 0x40, 0x40]), // 292 NEGBINOMDIST
    Some(&[0x40, 0x40, 0x40, 0x40]), // 293 NORMDIST
    Some(&[0x40]), // 294 NORMSDIST
    Some(&[0x40, 0x40, 0x40]), // 295 NORMINV
    Some(&[0x40]), // 296 NORMSINV
    Some(&[0x40, 0x40, 0x40]), // 297 STANDARDIZE
    Some(&[0x40]), // 298 ODD
    Some(&[0x40, 0x40]), // 299 PERMUT
    Some(&[0x40, 0x40, 0x40]), // 300 POISSON
    Some(&[0x40, 0x40, 0x40]), // 301 TDIST
    Some(&[0x40, 0x40, 0x40, 0x40]), // 302 WEIBULL
    Some(&[0x60, 0x60]), // 303 SUMXMY2
    Some(&[0x60, 0x60]), // 304 SUMX2MY2
    Some(&[0x60, 0x60]), // 305 SUMX2PY2
    Some(&[0x60, 0x60]), // 306 CHITEST
    Some(&[0x60, 0x60]), // 307 CORREL
    Some(&[0x60, 0x60]), // 308 COVAR
    Some(&[0x40, 0x60, 0x60]), // 309 FORECAST
    Some(&[0x60, 0x60]), // 310 FTEST
    Some(&[0x60, 0x60]), // 311 INTERCEPT
    Some(&[0x60, 0x60]), // 312 PEARSON
    Some(&[0x60, 0x60]), // 313 RSQ
    Some(&[0x60, 0x60]), // 314 STEYX
    Some(&[0x60, 0x60]), // 315 SLOPE
    Some(&[0x60, 0x60, 0x40, 0x40]), // 316 TTEST
    Some(&[0x60, 0x60, 0x40, 0x40]), // 317 PROB
    Some(&[0x20]), // 318 DEVSQ
    Some(&[0x20]), // 319 GEOMEAN
    Some(&[0x20]), // 320 HARMEAN
    Some(&[0x20]), // 321 SUMSQ
    Some(&[0x20]), // 322 KURT
    Some(&[0x20]), // 323 SKEW
    Some(&[0x20, 0x40, 0x40]), // 324 ZTEST
    Some(&[0x20, 0x40]), // 325 LARGE
    Some(&[0x20, 0x40]), // 326 SMALL
    Some(&[0x20, 0x40]), // 327 QUARTILE
    Some(&[0x20, 0x40]), // 328 PERCENTILE
    Some(&[0x20, 0x40, 0x40]), // 329 PERCENTRANK
    Some(&[0x60]), // 330 MODE
    Some(&[0x20, 0x40]), // 331 TRIMMEAN
    Some(&[0x40, 0x40]), // 332 TINV
    None, // 333
    None, // 334
    None, // 335
    Some(&[0x40]), // 336 CONCATENATE
    Some(&[0x40, 0x40]), // 337 POWER
    None, // 338
    None, // 339
    None, // 340
    None, // 341
    Some(&[0x40]), // 342 RADIANS
    Some(&[0x40]), // 343 DEGREES
    Some(&[0x40, 0x20]), // 344 SUBTOTAL
    Some(&[0x20, 0x40, 0x20]), // 345 SUMIF
    Some(&[0x20, 0x40]), // 346 COUNTIF
    Some(&[0x20]), // 347 COUNTBLANK
    None, // 348
    None, // 349
    Some(&[0x40, 0x40, 0x40, 0x40]), // 350 ISPMT
    Some(&[0x40, 0x40, 0x40]), // 351 DATEDIF
    Some(&[0x40]), // 352 DATESTRING
    Some(&[0x40, 0x40]), // 353 NUMBERSTRING
    Some(&[0x40, 0x40]), // 354 ROMAN
    None, // 355
    None, // 356
    None, // 357
    Some(&[0x40, 0x20]), // 358 GETPIVOTDATA
    Some(&[0x40, 0x40]), // 359 HYPERLINK
    Some(&[0x20]), // 360 PHONETIC
    Some(&[0x20]), // 361 AVERAGEA
    Some(&[0x20]), // 362 MAXA
    Some(&[0x20]), // 363 MINA
    Some(&[0x20]), // 364 STDEVPA
    Some(&[0x20]), // 365 VARPA
    Some(&[0x20]), // 366 STDEVA
    Some(&[0x20]), // 367 VARA
    None, // 368
];
