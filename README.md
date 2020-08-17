# POI 使用案例

##OLE 2文档的POIFS
> POIFS是POI中最古老，最稳定的部分。这是我们的OLE 2复合文档格式到纯Java的移植。它同时支持读写功能。根据定义，我们用于二进制（非XML）Microsoft Office格式的所有组件最终都依赖于它。请参阅POIFS项目页面以获取更多信息。

##HSSF和XSSF for Excel文档
> HSSF是我们将Microsoft Excel 97（-2003）文件格式（BIFF8）移植到纯Java的端口。XSSF是Microsoft Excel XML（2007+）文件格式（OOXML）到纯Java的移植。SS是一个使用通用API为两种格式提供通用支持的软件包。它们都支持读写功能。请参阅 HSSF + XSSF项目页面以获取更多信息。

##用于Word文档的HWPF和XWPF
> HWPF是我们将Microsoft Word 97（-2003）文件格式移植到纯Java的端口。它支持读取和有限的写入功能。它还为较旧的Word 6和Word 95格式提供了简单的文本提取支持。请参阅HWPF项目页面以获取更多信息。该组件仍处于开发的早期阶段。它已经可以读写简单的文件。
>我们还在根据OOXML规范为WordprocessingML（2007+）格式开发XWPF。这提供了对较简单文件的读写支持，以及文本提取功能。

##PowerPoint文档的HSLF和XSLF
> HSLF是我们将Microsoft PowerPoint 97（-2003）文件格式移植到纯Java的端口。它支持读写功能。请参阅HSLF项目页面以获取更多信息。
>我们还在根据OOXML规范为PresentationML（2007+）格式开发XSLF。

##HPSF OLE 2文档属性
> HPSF是OLE 2属性集格式到纯Java的移植。属性集通常用于存储文档的属性（标题，作者，最后修改日期等），但它们也可以用于特定于应用程序的目的。
> HPSF支持读取和写入属性。
> 请参阅HPSF项目页面以获取更多信息。

##用于Visio文档的HDGF和XDGF
> HDGF是我们将Microsoft Visio 97（-2003）文件格式移植到纯Java的端口。当前，它仅支持非常低的阅读水平，并且支持简单的文本提取。请参阅HDGF / Diagram项目页面以获取更多信息。
> XDGF是Microsoft Visio XML（.vsdx）文件格式到纯Java的移植。它比HDGF支持更多。请参阅XDGF / Diagram项目页面以获取更多信息。

##HPBF发布者文档
> HPBF是我们将Microsoft Publisher 98（-2007）文件格式移植到纯Java的端口。目前，它仅支持低水平读取大约一半的文件部分，并支持简单的文本提取。请参阅HPBF项目页面以获取更多信息。

##TNEF的HMEF（winmail.dat）Outlook附件
> HMEF是Microsoft TNEF（传输中性编码格式）文件格式到纯Java的移植。Outlook有时会使用TNEF对消息进行编码，通常会以winmail.dat的形式出现。HMEF当前仅支持较低级别的阅读，但我们希望添加文本和附件提取。请参阅HMEF项目页面以获取更多信息。

##HSMF for Outlook邮件
> HSMF是Microsoft Outlook消息文件格式到纯Java的移植。目前，它仅包含MSG文件的某些文本内容以及一些附件。进一步的支持和文档进展缓慢。目前，建议用户参考单元测试以供使用。请参阅HSMF项目页面以获取更多信息。
> Microsoft最近已将Outlook文件格式添加到其OSP中。现在可以获取更多信息，从而使实现此API更加容易。