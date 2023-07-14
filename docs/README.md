# Cosmic相关脚本使用文档

*上一次更新于 2023/07/09*

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

### 文档

#### 类 `class CosmicReqExcel(path: str)`

- 此类负责读取一个关于cosmic要求的单文件（非汇总表，为单cosmic评估发起的附件），一般为"cosmic软件发起"文件下的"附件5：..."
  
  - **`path`**: 该excel/csv文件的绝对或相对路径
    
    - **注意：路径之间的间隔符要为`\\`或`/`, 请在运行前确认**

- **方法 `load_excel() -> None`**
  
  - 读取在指定路径下的excel文件，支持`xlsx`和`xls`两种文件格式。数据加载进来会自动转为`pandas.Dataframe`

- **方法 `load_csv() -> None`**
  
  - 读取在指定路径下的csv文件。数据加载进来会自动转为`pandas.Dataframe`

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

- **方法 `load_csv() -> None`**
  
  - 读取在指定路径下的csv文件。数据加载进来会自动转为`pandas.Dataframe`

- **方法 `set_sheet_name(sheet_name: str) -> None`**
  
  - 设置需求汇总表中想要处理sheet名称，因为需求汇总表里可能有很多个小汇总表。**当您已经实例化此类后并想处理另一个小汇总表则可以通过这种方法设置**

- **方法 `print_df() -> None`**
  
  - 打印加载进来的每一页Excel表格到终端。格式为`pandas.Dataframe`的格式

- **方法 `print_df_specific() -> None`**
  
  - 打印通过`sheet_name`设置好的那一页表格到终端。格式为`pandas.Dataframe`的格式

- **方法 `check_ratio() -> list[str, None]`**
  
  - 检查需求汇总表里指定的页（sheet）的Cosmic送审工作量和Cosmic送审功能点之间的关系，并返回一个不符合**0.79**比例的所有条目的列表。这个比例可以手动设置。

- **方法 `check_file(req_num: int) -> dict`**
  
  - 检查需求汇总表里指定的页中的一个条目（行）和它所对应的文件夹。返回相应结果。
  
  - **`req_num`**: 该需求序号
  
  - 检查的项为：
    
    - 该汇总表中是否不存在该需求序号
    
    - 该汇总表中是否存在重复的该需求序号
    
    - 该总文件夹下是否存在该需求序号的子文件夹
    
    - 该需求"是否适用cosmic"列是否和excel/csv文件数量是否匹配，并且文件名开头的附件号码是否正确
    
    - 该excel是否可以被正常读取
    
    - 该需求名称和文件中的需求文件名称是否匹配
    
    - 该需求cosmic送审功能点是否和文件中的CFP总和匹配

- **方法 `check_all_files() -> dict[str, list[dict, None]]`**
  
  - 检查需求汇总表里指定的页中的所有条目（行）和它们所各自对应的文件夹。返回一个汇总所有结果和该方法总花费时间的字典。由于此方法的返回较为复杂，以下是返回的汇总字典格式范例
    
    ```json
    {
        "results": [{
            result 1
        }, {
            ...
        }],
        "time": float
    }
    ```
