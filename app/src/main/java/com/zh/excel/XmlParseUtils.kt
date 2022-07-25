package com.zh.excel

import android.util.Log
import org.w3c.dom.Element
import java.io.IOException
import java.io.InputStream
import javax.xml.parsers.DocumentBuilderFactory

object XmlParseUtils {

    private const val TAG = "XmlParseUtils"

    /**
     * 解析 Android 中 string.xml 为实体类
     *
     * @param inputStream 文件流
     * @return 实体类 List
     */
    fun dom(inputStream: InputStream?): ArrayList<StringModel> {
        val persons = ArrayList<StringModel>()
        try {
            val factory = DocumentBuilderFactory.newInstance()
            val builder = factory.newDocumentBuilder()
            val document = builder.parse(inputStream)
            val element = document.documentElement
            val items = element.getElementsByTagName("string")
            for (i in 0 until items.length) {
                val personNode = items.item(i) as Element
                val name = personNode.getAttribute("name")
                val value = personNode.firstChild.nodeValue ?: ""
                if (value.isEmpty()) {
                    Log.w(TAG, "dom: value is null, continue")
                    continue
                }
                Log.w(TAG, "dom: name:$name value:$value")
                persons.add(StringModel(name, value))
            }
        } catch (e: Exception) {
            e.printStackTrace()
            Log.e(TAG, "dom: e:${e.message}")
        } finally {
            try {
                inputStream?.close()
            } catch (e: IOException) {
                e.printStackTrace()
            }
        }
        return persons
    }

}