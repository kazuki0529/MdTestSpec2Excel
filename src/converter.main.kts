#!/usr/bin/env kotlin

//@file:DependsOn("net.sf.jett:jett-core:0.11.0")
@file:DependsOn("org.jxls:jxls-poi:2.12.0")
@file:DependsOn("com.vladsch.flexmark:flexmark-all:0.64.0")

import com.vladsch.flexmark.ast.BulletList
import com.vladsch.flexmark.ast.FencedCodeBlock
import com.vladsch.flexmark.ast.Heading
import com.vladsch.flexmark.ast.OrderedList
import com.vladsch.flexmark.parser.Parser
import com.vladsch.flexmark.util.ast.Document
import com.vladsch.flexmark.util.ast.Node
import com.vladsch.flexmark.util.data.MutableDataSet
import org.jetbrains.kotlin.com.google.common.io.Files
import org.jxls.common.Context
import org.jxls.util.JxlsHelper
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.Paths

/**
 * テスト仕様エンティティ
 *
 * @property fileName
 * @property title
 * @property cases
 */
data class Spec(
    val fileName: String,
    val title: String,
    val cases: List<SpecCase>
)

/**
 * テストケースエンティティ
 *
 * @property mainItem
 * @property middleItem
 * @property smallItem
 * @property steps
 * @property expected
 * @property notes
 */
data class SpecCase(
    val mainItem: String,
    val middleItem: String,
    val smallItem: String,
    val steps: String,
    val expected: String,
    val notes: String
)

/**
 * Markdown形式のテスト仕様書のローダ
 *
 * @constructor
 * コンストラクタ
 *
 * @param specFile
 */
class SpecLoader(specFile: File) {
    var cursor: Node? = null
    private var document: Document

    var title = ""
        private set
    var mainItem = ""
        private set
    var middleItem = ""
        private set
    var smallItem = ""
        private set
    var steps = listOf<String>()
        private set
    var expected = listOf<String>()
        private set
    var notes = listOf<String>()
        private set

    var cases = mutableListOf<SpecCase>()
        private set

    private var case: MutableMap<String, String> = mutableMapOf()

    init {
        val parser = Parser.builder(MutableDataSet()).build()
        this.document = parser.parse(
            Files.readLines(specFile, StandardCharsets.UTF_8)
                .joinToString("\n")
        )
        this.cursor = this.document.firstChild
    }

    /**
     * 次のノードが存在するか確認する
     *
     * @return 次のノードが存在する場合 true を返し、存在しない場合は false を返す
     */
    fun hasNext(): Boolean = this.cursor?.next !== null

    /**
     * 次のノードに移動する
     * テストケースが終了していれば、合わせてテストケースを追加する
     *
     * @return 次のノードが存在する場合 true を返し、存在しない場合は false を返す
     */
    fun next(): Boolean {
        this.cursor = this.cursor?.next
        // 新しいケースの出現か、ファイルの末端か
        if (this.isNewCase() || this.cursor == null) {
            this.cases.add(
                SpecCase(
                    this.mainItem,
                    this.middleItem,
                    this.smallItem,
                    this.steps.joinToString("\n"),
                    this.expected.joinToString("\n"),
                    this.notes.joinToString("\n")
                )
            )
            this.notifyNewCase()
        }
        return this.cursor != null
    }

    /**
     * ノードが新しいテストケースに移動したか判断する
     *
     * @return 新しいテストケースに移動していれば true、移動していなければ false を返す
     */
    private fun isNewCase(): Boolean = this.cursor is Heading && this.steps.isNotEmpty() && this.expected.isNotEmpty()

    /**
     * 新しいケースの通知を受ける
     *
     * 新しいケース用にフィールドを初期化する
     */
    private fun notifyNewCase() {
        this.mainItem = ""
        this.middleItem = ""
        this.smallItem = ""
        this.steps = listOf()
        this.expected = listOf()
        this.notes = listOf()
    }

    /**
     * タイトルの通知を受ける
     *
     * @param value タイトル
     */
    fun notifyTitle(value: String) {
        this.title = value
    }

    /**
     * 大項目読み込みの通知を受ける
     *
     * @param value タイトル
     */
    fun notifyMainItem(value: String) {
        this.mainItem = value
    }

    /**
     * 中項目読み込みの通知を受ける
     *
     * @param value 中項目
     */
    fun notifyMiddleItem(value: String) {
        this.middleItem = value
    }

    /**
     * 小項目読み込みの通知を受ける
     *
     * @param value 小項目
     */
    fun notifySmallItem(value: String) {
        this.smallItem = value
    }

    /**
     * 確認手順読み込みの通知を受ける
     *
     * @param value 確認手順
     */
    fun notifySteps(value: String) {
        var num = 1
        this.steps = value.trim().split("\n").map {
            """${num++}. """ + it.replace("""^\d+\. """.toRegex(), "")
        }
    }

    /**
     * 確認項目読み込みの通知を受ける
     *
     * @param value 確認項目
     */
    fun notifyExpected(value: String) {
        this.expected = value.trim().split("\n").map {
            "・" + it.replace("""^[\*\+\-] \[ \] """.toRegex(), "")
        }
    }

    /**
     * 備考読み込みの通知を受ける
     *
     * @param value 備考
     */
    fun notifyNotes(value: String) {
        this.notes = value.split("\n").filter { !it.startsWith("```") }
    }
}

// 引数チェック
if (args.size > 3)
    throw IllegalArgumentException("converter.main.kts [path-to-markdown-dir] [path-to-template-excel-file] [path-to-output-excel-file]")

val mdSpecDir = args[0]
val template = args[1]
val out = args[2]


println("######### Start #########")

// Markdownの読み込み処理
val list: List<String> = File(mdSpecDir).list()?.filter { it.lowercase().endsWith(".md") } ?: listOf()
val specList = mutableListOf<Spec>()

list.forEach {
    println("""---------- $it ----------""")
    val file = Paths.get(mdSpecDir).resolve(it).toFile()
    val loader = SpecLoader(file)

    do {
        when (val current = loader.cursor) {
            // ヘッダ(h1, h2, h2, h3)
            // タイトルと大、中、小項目
            is Heading -> {
                when (current.level) {
                    1 -> {
                        loader.notifyTitle(current.text.toString())
                    }

                    2 -> {
                        loader.notifyMainItem(current.text.toString())
                    }

                    3 -> {
                        loader.notifyMiddleItem(current.text.toString())
                    }

                    4 -> {
                        loader.notifySmallItem(current.text.toString())
                    }
                }
            }

            // 番号付きリスト
            // チェック条件
            is OrderedList -> {
                loader.notifySteps(current.chars.toString())
            }

            // 箇条書き
            // 確認項目
            is BulletList -> {
                loader.notifyExpected(current.chars.toString())
            }

            // コードブロック
            // 備考
            is FencedCodeBlock -> {
                loader.notifyNotes(current.chars.toString())
            }
        }

    } while (loader.next())

    val spec = Spec(
        file.nameWithoutExtension,
        loader.title,
        loader.cases
    )
    specList.add(spec)
    println(spec)
}

// Excelテンプレート出力
val context = Context().apply {
    specList.forEach {
        this.putVar(it.fileName, it)
    }
}
JxlsHelper.getInstance().processTemplate(FileInputStream(template), FileOutputStream(out), context)

println("######### Finished #########")
