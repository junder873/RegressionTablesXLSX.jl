module RegressionTablesXLSX

using Reexport
@reexport using RegressionTables
using PythonCall

export XlsxTable, AbstractXlsx

include("xlsx.jl")

end # module RegressionTablesXLSX
