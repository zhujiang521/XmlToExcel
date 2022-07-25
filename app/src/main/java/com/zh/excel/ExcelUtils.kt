package com.zh.excel

import android.content.Context
import android.net.Uri
import android.widget.Toast
import jxl.Workbook
import jxl.WorkbookSettings
import jxl.format.Alignment
import jxl.format.Border
import jxl.format.BorderLineStyle
import jxl.format.Colour
import jxl.write.*
import java.io.*


/**
 * @author dmrfcoder
 * @date 2018/8/9
 */
object ExcelUtils {
    private var arial14format: WritableCellFormat? = null
    private var arial10format: WritableCellFormat? = null
    private var arial12format: WritableCellFormat? = null

    /**
     * 单元格的格式设置 字体大小 颜色 对齐方式、背景颜色等...
     */
    private fun format() {
        try {
            val arial14font = WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD)
            arial14font.colour = Colour.LIGHT_BLUE
            arial14format = WritableCellFormat(arial14font)
            arial14format?.alignment = Alignment.CENTRE
            arial14format?.setBorder(Border.ALL, BorderLineStyle.THIN)
            arial14format?.setBackground(Colour.VERY_LIGHT_YELLOW)
            val arial10font = WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD)
            arial10format = WritableCellFormat(arial10font)
            arial10format?.alignment = Alignment.CENTRE
            arial10format?.setBorder(Border.ALL, BorderLineStyle.THIN)
            arial10format?.setBackground(Colour.GRAY_25)
            val arial12font = WritableFont(WritableFont.ARIAL, 10)
            arial12format = WritableCellFormat(arial12font)
            //对齐格式
            arial10format?.alignment = Alignment.CENTRE
            //设置边框
            arial12format!!.setBorder(Border.ALL, BorderLineStyle.THIN)
        } catch (e: WriteException) {
            e.printStackTrace()
        }
    }


    /**
     * 初始化Excel
     *
     * @param fileName 导出excel存放的地址（目录）
     * @param colName  excel中包含的列名（可以有多个）
     */
    fun initExcel(fileName: String?, sheetName: String?, colName: Array<String>) {
        format()
        var workbook: WritableWorkbook? = null
        try {
            val file = File(fileName)
            if (!file.exists()) {
                file.createNewFile()
            }
            val workbookSettings = WorkbookSettings()
            workbookSettings.encoding = "UTF-8"
            workbookSettings.formulaAdjust = true
            workbookSettings.refreshAll = true
            val os = FileOutputStream(file)
            workbook = Workbook.createWorkbook(os)
            //设置表格的名字
            val sheet = workbook!!.createSheet(sheetName, 0)
            //创建标题栏
            sheet.addCell(Label(0, 0, "没事请关注公众号：江江安卓", arial14format) as WritableCell)
            for (col in colName.indices) {
                // Label(x,y,z) 代表单元格的第x+1列，第y+1行, 内容z
                // 在Label对象的子对象中指明单元格的位置和内容
                sheet.addCell(Label(col, 1, colName[col],  /*getHeader()*/arial10format))
                sheet.setColumnView(0, colName[col].length * 8)
            }
            //设置行高
            sheet.setRowView(0, 340)
            workbook.write()
        } catch (e: java.lang.Exception) {
            e.printStackTrace()
        } finally {
            if (workbook != null) {
                try {
                    workbook.close()
                } catch (e: java.lang.Exception) {
                    e.printStackTrace()
                }
            }
        }
    }


    /**
     * 写入实体类数据到Excel文件
     */
    fun <T> writeObjListToExcel(objList: List<T>?, fileName: String?, c: Context?): Uri? {
        if (objList != null && objList.isNotEmpty()) {
            var writebook: WritableWorkbook? = null
            var inputStream: InputStream? = null
            val file = File(fileName)
            try {
                inputStream = FileInputStream(file)
                val workbook = Workbook.getWorkbook(inputStream)
                writebook = Workbook.createWorkbook(file, workbook)
                val sheet = writebook.getSheet(0)
                if (objList.isNotEmpty()) {
                    for (j in objList.indices) {
                        //要写入表格的数据对象
                        val faceTestEntity: StringModel = objList[j] as StringModel
                        val list: MutableList<String> = ArrayList()
                        list.add(faceTestEntity.name)
                        list.add(faceTestEntity.value)
                        for (i in list.indices) {
                            sheet.addCell(Label(i, j + 2, list[i], arial12format))
                            if (list[i].length <= 4) {
                                //设置列宽 ，第一个参数是列的索引，第二个参数是列宽
                                sheet.setColumnView(i, list[i].length * 7)
                            } else {
                                //设置列宽
                                sheet.setColumnView(i, list[i].length * 4)
                            }
                        }
                        //设置行高，第一个参数是行数，第二个参数是行高
                        sheet.setRowView(j, 340)
                    }
                }
                writebook.write()
                c?.apply {
                    Toast.makeText(c, "导出Excel成功", Toast.LENGTH_SHORT).show()
                    return FileStorageUtils.getFileUri(context = c, file)
                }
            } catch (e: java.lang.Exception) {
                e.printStackTrace()
            } finally {
                if (writebook != null) {
                    try {
                        writebook.close()
                    } catch (e: java.lang.Exception) {
                        e.printStackTrace()
                    }
                }
                if (inputStream != null) {
                    try {
                        inputStream.close()
                    } catch (e: IOException) {
                        e.printStackTrace()
                    }
                }
            }
        }
        return null
    }

}