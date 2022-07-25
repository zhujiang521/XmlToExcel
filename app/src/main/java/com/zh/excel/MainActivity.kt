package com.zh.excel

import android.Manifest
import android.annotation.SuppressLint
import android.app.Activity
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.ContextCompat
import com.zh.excel.FileStorageUtils.getStorageFilePath
import java.io.File


class MainActivity : AppCompatActivity(), View.OnClickListener {

    companion object {
        const val PICK_FILE = 1
    }

    private lateinit var mainBtnSelect: Button
    private lateinit var mainTvFile: TextView
    private lateinit var mainEtName: EditText
    private lateinit var mainEtSheet: EditText
    private lateinit var mainBtnParse: Button
    private val stringList = ArrayList<StringModel>()

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        initView()
    }

    private fun initView() {
        mainBtnSelect = findViewById(R.id.mainBtnSelect)
        mainTvFile = findViewById(R.id.mainTvFile)
        mainEtName = findViewById(R.id.mainEtName)
        mainEtSheet = findViewById(R.id.mainEtSheet)
        mainBtnParse = findViewById(R.id.mainBtnParse)
        mainBtnSelect.setOnClickListener(this)
        mainBtnParse.setOnClickListener(this)
    }

    override fun onClick(v: View?) {
        when (v?.id) {
            R.id.mainBtnSelect -> {
                getFile()
            }
            R.id.mainBtnParse -> {
                toExcel()
            }
        }
    }

    @SuppressLint("QueryPermissionsNeeded")
    private fun toExcel() {
        if (stringList.isEmpty()) {
            Toast.makeText(this, "当前数据没有选择成功", Toast.LENGTH_SHORT).show()
            return
        }

        val dirPath = getStorageFilePath(this) + File.separator + "AndroidExcel"
        val file = File(dirPath)
        if (!file.exists()) file.mkdirs()
        val fileName = if (mainEtName.text.isNotEmpty()) "${
            mainEtName.text.toString().trim()
        }.xls" else "TestExcel.xls"
        val testExcelPath = dirPath + File.separator + fileName

        val sheetName =
            if (mainEtSheet.text.isNotEmpty()) mainEtSheet.text.toString().trim() else "测试表1"

        ExcelUtils.initExcel(testExcelPath, sheetName, arrayOf("string", "value"))

        val tempUri = ExcelUtils.writeObjListToExcel(stringList, testExcelPath, this);

        if (tempUri != null) {
            val intent = Intent(Intent.ACTION_SEND).apply {
                type = "*/*"
                putExtra(Intent.EXTRA_STREAM, tempUri)
                flags = Intent.FLAG_ACTIVITY_NEW_TASK
            }
            if (intent.resolveActivity(packageManager) != null) {
                startActivity(Intent.createChooser(intent, title))
                return
            }
        } else {
            Toast.makeText(this, "分享失败", Toast.LENGTH_SHORT).show()
        }
    }

    private fun getFile() {
        if (ContextCompat.checkSelfPermission(
                this,
                Manifest.permission.READ_EXTERNAL_STORAGE
            ) == PackageManager.PERMISSION_GRANTED
        ) {
            pickFile()
        } else {
            Toast.makeText(this, "麻烦手动打开此权限", Toast.LENGTH_SHORT).show()
        }
    }

    private fun pickFile() {
        val intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
        intent.addCategory(Intent.CATEGORY_OPENABLE)
        intent.type = "*/*"
        startActivityForResult(intent, PICK_FILE)
    }

    @Deprecated("Deprecated in Java")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        when (requestCode) {
            PICK_FILE -> {
                if (resultCode == Activity.RESULT_OK && data != null) {
                    val uri = data.data
                    if (uri != null) {
                        mainTvFile.text = "Uri:$uri"
                        val inputStream = contentResolver.openInputStream(uri)
                        // 执行文件读取操作
                        stringList.clear()
                        stringList.addAll(XmlParseUtils.dom(inputStream))
                    }
                }
            }
        }
    }


}