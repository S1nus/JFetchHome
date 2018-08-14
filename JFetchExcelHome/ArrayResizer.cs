using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace JFetchExcelHome {
	class ArrayResizer {

		static Queue<ExcelReference> ResizeJobs = new Queue<ExcelReference>();

		public static object Resize(object[,] array) {
			ExcelReference caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
			return Resize(array, caller);
		}

		public static object Resize(object[,] array, ExcelReference caller) {
			if (caller == null) {
				return array;
			}

			int rows = array.GetLength(0);
			int collumns = array.GetLength(1);

			if ((caller.RowLast - caller.RowFirst + 1 != rows) || 
				(caller.ColumnLast - caller.ColumnFirst + 1 != collumns)) {
				EnqueueResize(caller, rows, collumns);
				ExcelAsyncUtil.QueueAsMacro(DoResizing);
			}

			return array;
		}

		static void EnqueueResize(ExcelReference caller, int rows, int columns) {
			ExcelReference target = new ExcelReference(caller.RowFirst, caller.RowFirst + rows - 1, caller.ColumnFirst, caller.ColumnFirst + columns - 1, caller.SheetId);
			ResizeJobs.Enqueue(target);
		}

		static void DoResizing() {
			while (ResizeJobs.Count > 0) {
				DoResize(ResizeJobs.Dequeue());
			}
		}

		static void DoResize(ExcelReference target) {
			object oldEcho = XlCall.Excel(XlCall.xlfGetWorkspace, 40);
			object oldCalculationMode = XlCall.Excel(XlCall.xlfGetDocument, 14);

			try {
				XlCall.Excel(XlCall.xlcEcho, false);
				XlCall.Excel(XlCall.xlcOptionsCalculation, 3);

				string formula = (string)XlCall.Excel(XlCall.xlfGetCell, 41, target);

				ExcelReference firstCell = new ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId);

				bool isFormulaArray = (bool)XlCall.Excel(XlCall.xlfGetCell, 49, target);
				if (isFormulaArray) {
					object oldSelectionOnActiveSheet = XlCall.Excel(XlCall.xlfGetCell, 49, target);
					object oldActiveCell = XlCall.Excel(XlCall.xlfActiveCell);

					string firstCellSheet = (string)XlCall.Excel(XlCall.xlSheetNm, firstCell);
					XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { firstCellSheet });
                    object oldSelectionOnArraySheet = XlCall.Excel(XlCall.xlfSelection);
					XlCall.Excel(XlCall.xlcFormulaGoto, firstCell);

                    XlCall.Excel(XlCall.xlcSelectSpecial, 6);
                    ExcelReference oldArray = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);

                    oldArray.SetValue(ExcelEmpty.Value);
                    XlCall.Excel(XlCall.xlcSelect, oldSelectionOnArraySheet);
					XlCall.Excel(XlCall.xlcFormulaGoto, oldSelectionOnActiveSheet);
				}

				bool isR1C1Mode = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
				string formulaR1C1 = formula;
				if (!isR1C1Mode) {
					formulaR1C1 = (string)XlCall.Excel(XlCall.xlfFormulaConvert, formula, true, false, ExcelMissing.Value, firstCell);
				}
				object ignoredResult;
				XlCall.XlReturn retval = XlCall.TryExcel(XlCall.xlcFormulaArray, out ignoredResult, formulaR1C1, target);
				
				if (retval != XlCall.XlReturn.XlReturnSuccess) {
					firstCell.SetValue("'" + formula);
				}
			}
			finally {
				XlCall.Excel(XlCall.xlcEcho, oldEcho);
				XlCall.Excel(XlCall.xlcOptionsCalculation, oldCalculationMode);
			}
		}

	}
}
