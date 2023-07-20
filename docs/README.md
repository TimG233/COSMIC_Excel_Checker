# Cosmic相关脚本使用文档

*上一次更新于 2023/07/20*

### 快速开始

- Cosmic_Excel_Checker集合了一系列的相关脚本功能到一个工具包模块里(module), 首先需要安装Cosmic_Excel_Checker作为可调用模块并在您的代码中进行方法调用。

- 安装方法: PyPI?Github?

- **请注意：此模块要求python版本>=3.9**

- 安装完成后, 在自己的python代码文件中, 可以采用绝对或相对import来引入模块
  
  ```python
  import cosmicexcelchecker  # absolute import
  from cosmicexcelchecker import CosmicReqExcel
  ```

- 例如我想知道一个符合cosmic需求的（附件5）的总COSMIC功能点总和，就可以通过如下代码来实现，
  
  ```python
  from cosmicexcelchecker import CosmicReqExcel
  
  # Instantiate class
  cre = CosmicReqExcel(path='path/to/the/excel.xlsx')  # don't forget to replace path
  cre.load_excel()  # this is mandatory
  print(cre.get_CFP_total())  # print total CFP pts to the terminal
  ```

- 以上例子不难发现，`load_excel()`是非常关键的一步，任何关于excel/csv相关文件的类(class)都会有`load_excel()`或 `load_csv()`的方法。加载文件并非自动的原因是为了避免因用户创建太多实例而自动加载文件导致占用太多系统资源。**因此，您必须手动使用这个方法来加载excel/csv文件**

- 这一部分只是快速开始，至于所有类和其相关方法，将在下面文档详细介绍。

### 简介

- Cosmic Excel Checker的编写初衷是为了用程序来替代一些对于COSMIC相关文件大量且重复人工工作，由程序来完成。优点是速度快，出错概率低，且自由度比较高。除此之外，此模块使用较为简单，可快速上手使用。

- 而且，此模块的速度也是较快的，通过读取excel/csv相关文件并转化成pandas的dataframe并进行处理，相关速度将比直接用python默认方法的处理速度更快，尤其是在面对非常大的excel/csv文件时。

### 性能

- 对于`check_all_files()`的性能对比 (`check_all_files()`调用`check_file()`)，因为检查子过程描述填充颜色和CFP对比的这一项十分缓慢，故而对比了包含与不包含这一项的性能。
  
  |             | 不包含子过程描述填充检查 | 包含子过程描述填充检查 |
  | ----------- | ------------ | ----------- |
  | 周期 (cycles) | 100          | 100         |
  | 最慢单次(秒)     | 6.77221      | 31.73421    |
  | 最快单次(秒)     | 3.29799      | 17.5083     |
  | 平均单次(秒)     | 3.88253      | 20.90706    |
  | 总评估访问       | 5400         | 5400        |
  | 单评估平均耗时     | 0.0719       | 0.38717     |

### 文档

#### 类 `class CosmicReqExcel(path: str)`

- 此类负责读取一个关于cosmic要求的单文件（非汇总表，为单cosmic评估发起的附件），一般为"cosmic软件发起"文件下的"附件5：..."
  
  - **`path`**: 该excel/csv文件的绝对或相对路径
    
    - **注意：路径之间的间隔符要为`\\`或`/`, 请在运行前确认**

- **方法 `load_excel() -> None`**
  
  - 读取在指定路径下的excel文件，支持`xlsx`和`xls`两种文件格式。数据加载进来会自动转为`pandas.Dataframe`

- **方法 `print_df() -> None`**
  
  - 打印加载进来的每一页Excel表格到终端。格式为`pandas.Dataframe`的格式

- **方法 `get_req_name() -> Union[str, None]`**
  
  - 获取需求名称（OPEX-需求名称），会返回一个名称的字符串或None（若缺失评估模板或没有需求名称的话）

- **方法 `get_CFP_total() -> Union[float, None]`**
  
  - 获取该需求分析的所有cfp点总和。会返回总CFP点的浮点数或None（若缺失评估模板或丢失CFP点数列）。**CFP列空白行默认为0**

- **方法 `check_CFP_column() -> dict`**
  
  - 对比CFP和子过程列，返回匹配的结果。可以通过此方法检查CFP列是否存在空行

#### 类 `class FindExcels()`

- 此类负责寻找一个目录下所有的excel/csv文件。**请注意，为了方便实用，此类所有方法都为静态方法，代表着您不需要实例化此类而是可以直接调用类中的静态方法**

- **静态方法 `path_format(path: str) -> str`**
  
  - 会替换所有路径分隔符`\\`为`/`, 以方便程序处理。
  
  - **`path`**: 指定的绝对/相对路径

- **静态方法 `find_excels(path: str) -> list[str, None]`**
  
  - 寻找指定路径下所有excel/csv文件，并将符合条件的文件路径存储到一个列表(list)中，若没有符合条件的文件，则会返回一个空列表。 
    
    - 注意：这个静态方法会在过程中调用`path_format`静态方法，所以您如果只想寻找所有符合条件的文件，则不需要亲自调用`path_format`方法
  
  - **`path`**: 指定的绝对/相对路径

#### 类 `class ResultSummary(path: str, folders_path: str, sheet_name: str)`

- 此类负责加载一个cosmic需求汇总表的excel/csv文件并进行相关操作。
  
  - **`path`**: 该需求汇总表的绝对/相对路径
  
  - **`folders_path`**: 要对照的所有需求文件夹所在的总文件夹路径
  
  - **`sheet_name`**: 需求汇总表中想要处理的sheet名称，因为需求汇总表里可能有很多个小汇总表。
  
  - **请注意：此类实例化支持`path`和`folders_path`两个参数，这说明汇总表和所有需求文件夹可以在计算机的不同位置。但如果在不同位置的话，请确定提供的这两个路径都为绝对路径**

- **方法 `load_excel() -> None`**
  
  - 读取在指定路径下的excel文件，支持`xlsx`和`xls`两种文件格式。数据加载进来会自动转为`pandas.Dataframe`

- **方法 `set_sheet_name(sheet_name: str) -> None`**
  
  - 设置需求汇总表中想要处理sheet名称，因为需求汇总表里可能有很多个小汇总表。**当您已经实例化此类后并想处理另一个小汇总表则可以通过这种方法设置**

- **方法 `print_df() -> None`**
  
  - 打印加载进来的每一页Excel表格到终端。格式为`pandas.Dataframe`的格式

- **方法 `print_df_specific() -> None`**
  
  - 打印通过`sheet_name`设置好的那一页表格到终端。格式为`pandas.Dataframe`的格式

- **方法 `check_ratio() -> list[str, None]`**
  
  - 检查需求汇总表里指定的页（sheet）的Cosmic送审工作量和Cosmic送审功能点之间的关系，并返回一个不符合**0.79**比例的所有条目的列表。这个比例可以手动设置。

- **方法 `check_file(req_num: int, check_final_confirmation: bool = True, check_highlight_cfp: bool = True) -> dict`**
  
  - 检查需求汇总表里指定的页中的一个条目（行）和它所对应的文件夹。返回相应结果。
  
  - **`req_num`**: 该需求序号
  
  - **`check_final_confirmation`**: 是否检查结算评估确认表相关信息
  
  - **`check_highlight_cfp`**: 是否检查子过程描述高亮和对应cfp点关系是否正确
  
  - 检查的项为：
    
    - 该汇总表中是否不存在该需求序号
    
    - 该汇总表中是否存在重复的该需求序号
    
    - 该总文件夹下是否存在该需求序号的子文件夹
    
    - 该需求"是否适用cosmic"列是否和excel/csv文件数量是否匹配，并且文件名开头的附件号码是否正确
    
    - 该excel是否可以被正常读取
    
    - 该需求名称和文件中的需求文件名称是否匹配
    
    - 该需求cosmic送审功能点是否和文件中的CFP总和匹配
    
    - 该需求文件中CFP总和是否和系数表匹配（可选择）
    
    - 该需求每一个子过程描述填充颜色是否和CFP点匹配（可选择，耗时长）

- **方法 `check_all_files(check_final_confirmation: bool = True, check_highlight_cfp: bool = True) -> dict[str, list[dict, None]]`**
  
  - 检查需求汇总表里指定的页中的所有条目（行）和它们所各自对应的文件夹。返回一个汇总所有结果和该方法总花费时间的字典。由于此方法的返回较为复杂，以下是返回的汇总字典格式范例
    
    ```json
    {
        "results": [{
            /* result 1 */
        }, {
            /* ... */
        }],
        "time": /* float */
    }
    ```
  
  - **`check_final_confirmation`**: 是否检查结算评估确认表相关信息
  
  - **`check_highlight_cfp`**: 是否检查子过程描述高亮和对应cfp点关系是否正确

#### 类 `class CheckObf()`

- 此类负责对比判断两个字符串的编辑距离，并且使用比例来判断两个字符串是否为相似字符串。
  
  - **请注意：字符串的相近并不能代表二者语义相近**，例：”我每天都喜欢吃早饭”和“我每天都**不**喜欢吃早饭”。二者非常相似（编辑距离仅为1），但语义却截然不同。
  
  - **若要判断语义，请考虑训练NLP模型来处理这类问题**

- **静态方法 `compare(string1: str, string2: str) -> int`**
  
  - 检查两个给定字符串并且计算出编辑距离，并返回整数距离结果。
  
  - **注意**：本方法采取Levenshtein Distance的算法，在动态规划下时间复杂度仍然很高: $O(mn)$ ($m, n$为两个字符串长度)。若对比两个超过2500字符长的字符串，请考虑使用其他方法。
  
  - 性能概览 (对比字符为固定长度随机生成，小于1纳秒则显示为0)
    
    | 对比字符长度 | 所需时间（秒）     |
    | ------ | ----------- |
    | 10     | 0           |
    | 20     | 0           |
    | 50     | 0.0035936   |
    | 100    | 0.012002    |
    | 500    | 0.2947539   |
    | 1000   | 1.639346    |
    | 2500   | 9.9067159   |
    | 5000   | 24.2416176  |
    | 10000  | 104.8854004 |

- **静态方法 `similarity(string1: str, string2: str, ratio: Union[float, None] = None) -> Union[float, bool]`**
  
  - 通过`compare(string1, string2)`方法来获取编辑距离，对比两个字符串中的较长者来对比比率，若不提供ratio则返回为浮点数的比率。若提供ratio且编辑距离大于ratio则判定为相似，返回布尔值。
    
    - 对比公式：`(len(string_longer) - edit_distance) / len(string_longer)`

#### 关于Conf

- **方法 `set_config(config: dict) -> None`**
  
  - 传入一个字典来修改单个/多个`conf.py`中的固定值。
  
  - 示例（添加，更改）
    
    ```py
    from conf import set_config, SR_FINAL_CONFIRMATION
    
    # **注意**：字典dict不允许一个字典中含有多个相同的key
    set_config(
        {
            'CFP_SHEET_NAMES': ['新表名1', '新表名2'],  # 移除列表类型旧值，填入新值
            'SR_FINAL_CONFIRMATION': SR_FINAL_CONFIRMATION.extend(['额外名称1', '额外名称2']),  # 添加多个额外值， 添加单值也可用.append(<value>)
            'CFP_COLUMN_NAME': 'CFP点数'  # 修改字符串类型固定值
            'WORKLOAD_CFP_RATIO': 0.85  # 修改整数/浮点数类型固定值
        }
    )
    ```
  
  - 具体固定值详解如下（传入值需要严格按照各个固定值类型来传入）
    
    | 固定值名称                             | 类型     | 默认值                               | 注释                                           |
    | --------------------------------- | ------ | --------------------------------- | -------------------------------------------- |
    | `CFP_SHEET_NAMES`                 | `list` | `['功能点拆分表', 'COSMIC软件评估标准模板']`    | cosmic功能点拆分表的sheet名称                         |
    | `NONCFP_SHEET_NAMES`              | `str`  | `非COSMIC评估工作量填写说明`                | 非cosmic评估表详细信息sheet名称                        |
    | `CFP_COLUMN_NAME`                 | `str`  | `CFP`                             | cosmic功能拆分表中CFP数字点数列名称                       |
    | `SUB_PROCESS_NAME`                | `str`  | `子过程描述`                           | cosmic功能拆分表中子过程描述列名称                         |
    | `RS_SKIP_ROWS`                    | `int`  | `9`                               | 汇总表格中跳过录入前n行 (跳过表头的一些总结数据，只读每一个需求相关的名头及各行数据) |
    | `Workload_CFP_Ratio`              | `int`  | `0.79`                            | 送审工作量和CFP的比率，一般默认0.79标准                      |
    | `RS_WORKLOAD_NAME`                | `str`  | `cosmic送审工作量`                     | 汇总表格中cosmic送审工作量列的名称                         |
    | `RS_TOTAL_CFP_NAME`               | `str`  | `cosmic送审功能点`                     | 汇总表格中cosmic送审功能点列名称                          |
    | `RS_REQ_NUM`                      | `str`  | `需求序号`                            | 汇总表格中需求序号列名称                                 |
    | `RS_REQ_NAME`                     | `str`  | `实施需求名称`                          | 汇总表格中实施需求名称列名称                               |
    | `RS_QLF_COSMIC`                   | `str`  | `是否适用cosmic`                      | 汇总表格中是否适用cosmic列名称（是/否/混合型）                  |
    | `SR_COSMIC_REQ_NAME`              | `str`  | `OPEX-需求名称`                       | cosmic功能拆分表中需求名称列名称                          |
    | `SR_NONCOSMIC_REQ_NAME`           | `str`  | `需求名称`                            | 非cosmic功能拆分表中需求名称列名称                         |
    | `COEFFICIENT_SHEET_NAME`          | `str`  | `系数表`                             | cosmic功能拆分表中系数表sheet名称                       |
    | `COEFFICIENT_SHEET_DATA_COL_NAME` | `str`  | `数值`                              | cosmic功能拆分表中系数表sheet里数值列名称                   |
    | `SR_SUBFOLDER_NAME`               | `str`  | `COSMIC评估发起`                      | 单个需求所在的父级文件夹名称                               |
    | `SR_COSMIC_FILE_PREFIX`           | `str`  | `附件5`                             | 单个需求cosmic统计表文件名开头                           |
    | `SR_NONCOSMIC_FILE_PREFIX`        | `str`  | `附件4`                             | 单个需求非cosmic统计表文件名开头                          |
    | `SR_NONCOSMIC_REQ_NUM`            | `str`  | `需求序号`                            | 非cosmic需求文件里需求序号列名称                          |
    | `SR_NONCOSMIC_PROJECT_NAME`       | `str`  | `项目名称`                            | 非cosmic需求文件里项目名称列名称                          |
    | `SR_FINAL_CONFIRMATION`           | `list` | `['结算评估确认表', '结论评估确认表', '结论认同表']` | cosmic需求文件中结算评估确认表sheet名称                    |
    | `SR_AC_REQ_NUM`                   | `str`  | `需求工单号`                           | cosmic需求文件中结算评估确认表里的需求工单号列名称                 |
    | `SR_AC_REQ_NAME`                  | `str`  | `需求名称`                            | cosmic需求文件中结算评估确认表里的需求名称列名称                  |
    | `SR_AC_REPORT_NUM`                | `str`  | `上报工作量\n（人天）`                     | cosmic需求文件中结算评估确认表里的上报工作量列名称                 |
    | `SR_AC_FINAL_NUM`                 | `str`  | `最终结果\n（人天）`                      | cosmic需求文件中结算评估确认表里的最终结果列名称                  |
    | `SR_AC_FINAL_NUM_LIMIT`           | `int`  | `3.4`                             | cosmic需求文件中结算评估确认表最终人天上限                     |
