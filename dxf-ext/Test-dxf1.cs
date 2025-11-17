using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Text.Json;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using netDxf;
using ClosedXML.Excel;
using System.Globalization;
using System.Text.RegularExpressions; // Added for pattern matching

class MappingColumn { public string col { get; set; } public string attr { get; set; } public string source { get; set; } }
class Mapping { public List<MappingColumn> columns { get; set; } }
