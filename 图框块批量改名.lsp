;;; ======================================================================
;;; 图框块批量改名.lsp  (v8.5)
;;; ----------------------------------------------------------------------
;;; 功能 : 批量修改目标文件夹（可选含所有子文件夹）内 DWG/DXF 的图框块定义名
;;;        块名支持通配符匹配（如 *-S-A3-1），新块名必须手动输入
;;; 命令 : TKGM
;;; 原理 : 逐张 打开->重命名块->存盘->关闭（编辑器内 COM；ZWCAD 的
;;;        ObjectDBX 进程内创建会死锁，不可用，故不采用）
;;; 安全 : 修改前自动备份到 <根文件夹>\备份_日期时间\（保持相对路径）；
;;;        会话中已打开的图纸自动跳过；目标名已存在则跳过（可重复跑）；
;;;        单张出错只记失败，不中断批处理
;;; 适用 : 中望 ZWCAD / AutoCAD 2004+ 完整版；文件编码 ANSI(GBK)
;;; 用法 : APPLOAD 加载本文件 -> 命令行输入 TKGM
;;; ======================================================================

(vl-load-com)

(setq *tkxg:reg* "HKEY_CURRENT_USER\\Software\\TKGM")

;; 已打开图纸处理模式：T=只改名不代存盘（该图纸有未保存修改时）
(setq *tkxg:nosave* nil)

;; ---------- 工具函数 ------------------------------------------------

(defun tkxg:say (s) (princ (strcat "\n  " s)))

;; 全路径 -> 上级目录（统一反斜杠，不含末尾斜杠）
(defun tkxg:parentdir (p / i)
  (setq p (vl-string-translate "/" "\\" p))
  (if (setq i (vl-string-position 92 p nil t))   ; 92 = "\"
    (substr p 1 i)
    p
  )
)

;; 路径归一化：连续反斜杠压缩为单个
(defun tkxg:fixpath (p / i)
  (while (setq i (vl-string-search "\\\\" p))
    (setq p (strcat (substr p 1 (1+ i)) (substr p (+ i 3))))
  )
  p
)

;; 时间戳 yyyymmddhhmmss（手工删小数点，空目标串的 translate 不可靠）
(defun tkxg:stamp (/ s i)
  (setq s (rtos (getvar "CDATE") 2 8))
  (while (setq i (vl-string-search "." s))
    (setq s (strcat (substr s 1 i) (substr s (+ i 2))))
  )
  (substr s 1 14)
)

;; 字符串列表拼接 "a, b, c"
(defun tkxg:join (lst sep / s)
  (setq s "")
  (foreach x lst
    (setq s (if (= s "") x (strcat s sep x)))
  )
  s
)

;; 字符串分割
(defun tkxg:split (s sep / pos out)
  (setq out nil)
  (while (setq pos (vl-string-search sep s))
    (setq out (cons (substr s 1 pos) out)
          s   (substr s (+ pos (strlen sep) 1)))
  )
  (reverse (cons s out))
)

;; 逐级创建目录
(defun tkxg:mkdirs (path / acc seg)
  (setq path (vl-string-translate "/" "\\" path)
        acc  "")
  (foreach seg (tkxg:split path "\\")
    (if (/= seg "")
      (progn
        (setq acc (if (= acc "") seg (strcat acc "\\" seg)))
        (vl-mkdir acc)
      )
    )
  )
)

;; 收集 dir 下所有 DWG/DXF 的相对路径；rec=T 时递归子文件夹
;; （自动跳过 备份* 目录和 ~$ 临时文件；注意 wcmatch 的 ~ 是取反符，
;;   临时文件用字符串前缀判断，不能用 "~$*"）
(defun tkxg:collect (dir rec / f out)
  (setq out '())
  (foreach f (append (vl-directory-files dir "*.dwg" 1)
                     (vl-directory-files dir "*.dxf" 1))
    (if (/= (substr f 1 2) "~$") (setq out (cons f out)))
  )
  (setq out (reverse out))
  (if rec
    (foreach sub (vl-directory-files dir "*.*" -1)
      (if (and (/= sub ".") (/= sub "..")
               (/= (substr sub 1 2) "~$")
               (not (wcmatch sub "*备份*"))     ; 跳过 备份*/修正备份* 等
               (vl-file-directory-p (strcat dir "\\" sub)))
        (setq out (append out
                    (mapcar '(lambda (x) (strcat sub "\\" x))
                            (tkxg:collect (strcat dir "\\" sub) t))))
      )
    )
  )
  out
)

;; 文件是否已在本 CAD 会话中打开（任何内部出错都视为未打开）
(defun tkxg:isopen (fname / res fn)
  (setq res nil)
  (vl-catch-all-apply
    '(lambda ()
       (vlax-for d (vla-get-Documents (vlax-get-acad-object))
         (setq fn (vla-get-FullName d))
         (if (and (eq (type fn) 'STR) (= (strcase fname) (strcase fn)))
           (setq res t)
         )
       )
     )
  )
  res
)

;; 文件夹选择对话框（Windows Shell COM，与 CAD 无关）；取消/失败返回 nil
(defun tkxg:browsefolder (title / sh fldr item path)
  (setq path nil)
  (setq sh (vl-catch-all-apply 'vlax-create-object (list "Shell.Application")))
  (if (and sh (not (vl-catch-all-error-p sh)))
    (progn
      (setq fldr (vl-catch-all-apply 'vlax-invoke
                   (list sh 'BrowseForFolder 0 title 0)))
      (if (and fldr (not (vl-catch-all-error-p fldr)))
        (progn
          (setq item (vl-catch-all-apply 'vlax-get-property (list fldr 'Self)))
          (if (and item (not (vl-catch-all-error-p item)))
            (setq path (vl-catch-all-apply 'vlax-get-property (list item 'Path))))
          (if item (vl-catch-all-apply 'vlax-release-object (list item)))
        )
      )
      (if fldr (vl-catch-all-apply 'vlax-release-object (list fldr)))
      (vlax-release-object sh)
    )
  )
  ;; 虚拟文件夹（"::{...}"）、空串、不存在的路径一律视为取消
  (if (and (eq (type path) 'STR) (/= path "")
           (vl-file-directory-p path) (/= (substr path 1 2) "::"))
    path
    nil
  )
)

;; 读文件头判断落盘格式："DWG"(AC1x 二进制) / "DXF"(文本) / "?"
(defun tkxg:filekind (fname / fh c s)
  (setq s "")
  (if (setq fh (open fname "r"))
    (progn
      (repeat 4
        (setq c (read-char fh) s (strcat s (chr c)))
      )
      (close fh)
    )
  )
  (if (= s "AC10") "DWG" "DXF")
)

;; 工作包 XML 提取字段值（Base64 文本）
(defun tkxg:xmlget (fname tag / fh line a b val key1 key2)
  (setq val  nil
        key1 (strcat "<" tag)
        key2 (strcat "</" tag ">"))
  (if (setq fh (open fname "r"))
    (progn
      (while (and (not val) (setq line (read-line fh)))
        (if (or (vl-string-search (strcat key1 " ") line)
                (vl-string-search (strcat key1 ">") line))
          (progn
            (setq a (vl-string-search ">" line)
                  b (vl-string-search key2 line))
            (if (and a b (> b a))
              (setq val (substr line (+ a 2) (- b a 1)))
            )
          )
        )
      )
      (close fh)
    )
  )
  val
)

;; Base64 -> 字节表
(defun tkxg:b64bytes (s / tbl out i grp n1 n2 n3 n4 pad b1 b2 b3)
  (setq tbl "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        out nil
        i   1)
  (while (<= i (- (strlen s) 3))
    (setq grp (substr s i 4)
          i   (+ i 4)
          pad 0)
    (while (= (substr grp 4 1) "=")
      (setq grp (strcat (substr grp 1 3) "A") pad (1+ pad))
    )
    (setq n1 (vl-string-search (substr grp 1 1) tbl)
          n2 (vl-string-search (substr grp 2 1) tbl)
          n3 (vl-string-search (substr grp 3 1) tbl)
          n4 (vl-string-search (substr grp 4 1) tbl))
    (if (and n1 n2 n3 n4)
      (progn
        (setq b1 (logior (lsh n1 2) (lsh n2 -4))
              b2 (logior (lsh (logand n2 15) 4) (lsh n3 -2))
              b3 (logior (lsh (logand n3 3) 6) n4))
        (setq out (cons b1 out))
        (if (< pad 2) (setq out (cons b2 out)))
        (if (< pad 1) (setq out (cons b3 out)))
      )
    )
  )
  (reverse out)
)

;; 解出工作包项目编号（PROJECT_NUMBER，Base64）；失败返回 nil
(defun tkxg:getprojno (xmlfile / v bytes s)
  (setq s nil)
  (if (and xmlfile (setq v (tkxg:xmlget xmlfile "PROJECT_NUMBER")))
    (progn
      (setq bytes (tkxg:b64bytes v))
      (if (vl-some '(lambda (b) (> b 127)) bytes)
        (tkxg:say "项目编号含非 ASCII 字符，请手动输入新块名。")
        (setq s (vl-list->string bytes))
      )
    )
  )
  s
)

;; 旧块名 -> 新块名: 项目编号 + 从 -甲/-乙/-S- 起的尾段（与原图框一致）
(defun tkxg:genname (old pnum / i tail mark)
  (setq tail nil)
  (foreach mark '("-甲" "-乙")
    (if (and (not tail) (setq i (vl-string-search mark old)))
      (setq tail (substr old (1+ i)))
    )
  )
  (if (not tail)
    (if (setq i (vl-string-search "-S-" old))
      (setq tail (substr old (1+ i)))
    )
  )
  (if tail (strcat pnum tail) nil)
)

;; 定位当前文件夹附近的工作包 XML：root\工作包、上级\工作包、root 本层
;; 返回候选路径列表（去重）
(defun tkxg:findxml (root / dirs dir lst f cands)
  (setq cands nil
        dirs  (list (strcat root "\\工作包")
                    (strcat (tkxg:parentdir root) "\\工作包")
                    root))
  (foreach dir dirs
    (if (vl-file-directory-p dir)
      (foreach f (vl-directory-files dir "*.xml" 1)
        (setq f (strcat dir "\\" f))
        (if (not (member (strcase f) (mapcar 'strcase cands)))
          (setq cands (cons f cands))
        )
      )
    )
  )
  (reverse cands)
)

;; 在单个文档内重命名匹配块并存盘，返回 ("状态" "消息")
;; *tkxg:pnum* 非空时为自动模式：目标名 = 项目编号 + 各块旧名尾段
(defun tkxg:rename-doc (ndoc pat newname fname / blocks mats exist nm tgt isx pick
                                     i cand)
  ;; 用 vlax-get 通用访问取 Blocks/Name（ZWCAD 未注册 vla-get-Blocks）
  (setq blocks (vl-catch-all-apply 'vlax-get (list ndoc 'Blocks))
        mats   nil
        exist  nil)
  (if (vl-catch-all-error-p blocks) (setq blocks nil))
  ;; 第一遍：收集匹配块（自动模式逐块算目标名）+ 检查是否已是目标名
  (vlax-for blk blocks
    (setq nm  (vl-catch-all-apply 'vlax-get (list blk 'Name)))
    (if (vl-catch-all-error-p nm) (setq nm nil))
    (setq tgt (if (and *tkxg:pnum* (eq (type nm) 'STR))
                (tkxg:genname nm *tkxg:pnum*)
                newname))
    (if (and (eq (type nm) 'STR) (eq (type tgt) 'STR)
             (= (strcase nm) (strcase tgt)))
      (setq exist t)
    )
    (setq isx (vl-catch-all-apply 'vla-get-IsXRef (list blk)))
    (if (and (eq (type nm) 'STR)
             (eq (type tgt) 'STR)
             (/= (substr nm 1 1) "*")                          ; 排除匿名块/模型布局空间
             (or (vl-catch-all-error-p isx) (= :vlax-false isx))  ; 排除外部参照
             (wcmatch (strcase nm) (strcase pat)))
      (setq mats (cons (list nm blk tgt) mats))
    )
  )
  (setq mats (reverse mats))
  ;; 排除已等于自身目标名的块（重复运行时幂等）
  (setq mats (vl-remove-if
               '(lambda (x) (= (strcase (car x)) (strcase (caddr x))))
               mats))
  (cond
    ;; 没有可改的块：若存在已达目标名的块说明早已改过
    ((not mats)
     (if exist
       (list "跳过" (if *tkxg:pnum*
                      (strcat "已是项目编号目标名（编号 " *tkxg:pnum* "）")
                      (strcat "已是目标名 [" newname "]")))
       (list "跳过" "无匹配的图框块")
     )
    )
    ;; 多个候选：交互选择要改哪一个（0 = 本张跳过），各自显示生成/指定的目标名
    ((cdr mats)
     (princ (strcat "\n  该图匹配到 " (itoa (length mats)) " 个图框块，请选择要改的块："))
     (setq i 0)
     (foreach cand mats
       (setq i (1+ i))
       (princ (strcat "\n    [" (itoa i) "] " (car cand) "  ->  " (caddr cand)))
     )
     (setq pick (getint "\n  输入序号(0=本张跳过) <0>: "))
     (if (or (not pick) (<= pick 0) (> pick (length mats)))
       (list "跳过" "多块匹配，未选择，本张跳过")
       (tkxg:do-rename ndoc (nth (1- pick) mats) fname)
     )
    )
    (t (tkxg:do-rename ndoc (car mats) fname))
  )
)

;; 实际执行重命名+存盘，cand=(旧名 块对象 目标名)
(defun tkxg:do-rename (ndoc cand fname / oldname newname err serr ok note isdxf saved tp)
  (setq oldname (car cand)
        newname (caddr cand)
        isdxf   (= (strcase (vl-filename-extension fname)) ".DXF"))
  ;; 重命名（目标名被占用等错误就地转失败，不炸批处理）
  (setq err (vl-catch-all-apply 'vla-put-name (list (cadr cand) newname)))
  (cond
    ((vl-catch-all-error-p err)
     (list "失败" (strcat "重命名出错: " (vl-catch-all-error-message err))))
    (t
     (if *tkxg:nosave*
       ;; 图纸有未保存修改：只改名不代存，避免把未完成内容一起存掉
       (list "OK" (strcat oldname " -> " newname
                  "（图纸正打开且有未保存修改：块名已改未存盘，请手动保存）"))
       (progn
         ;; 存盘链：
         ;;  DWG: Save -> SaveAs 原路径
         ;;  DXF: Save 必败 -> SaveAs 必须带显式格式参数（否则弹格式对话框），
         ;;       按 2018/2013/2010/2007/2004/2000/R12 DXF 从新到旧逐个试
         (setq ok   t
               note "")
         (setq serr (vl-catch-all-apply 'vla-Save (list ndoc)))
         (if (vl-catch-all-error-p serr)
           (progn
             (setq saved nil)
             (if isdxf
               (foreach tp '(65 61 49 37 25 13 1)
                 (if (not saved)
                   (progn
                     (setq serr (vl-catch-all-apply 'vla-SaveAs
                                  (list ndoc fname tp)))
                     (if (not (vl-catch-all-error-p serr)) (setq saved t))
                   )
                 )
               )
               (progn
                 (setq serr (vl-catch-all-apply 'vla-SaveAs (list ndoc fname)))
                 (if (not (vl-catch-all-error-p serr)) (setq saved t))
               )
             )
             (cond
               ((not saved)
                (setq ok   nil
                      note (if isdxf
                             "存盘出错: SaveAs 各 DXF 格式均失败"
                             (strcat "存盘出错: " (vl-catch-all-error-message serr)))))
               ((= (tkxg:filekind fname) "DWG")
                (setq note "（注意:已被存为DWG格式）"))
               (t (setq note "（DXF格式已保留）"))
             )
           )
         )
         (if (not ok)
           (list "失败" note)
           (list "OK" (strcat oldname " -> " newname note))
         )
       )
     )
    )
  )
)

;; 打开单张图纸 -> 改名 -> 存盘 -> 关闭，返回 ("状态" "消息")
(defun tkxg:process-open (fname pat newname / docs r ndoc fn res)
  (setq docs (vla-get-Documents (vlax-get-acad-object))
        ndoc nil)
  ;; 打开图纸（同步）
  (setq r (vl-catch-all-apply 'vla-Open (list docs fname)))
  ;; 按 FullName 定位新文档（不依赖 Open 返回值）
  (vlax-for d docs
    (setq fn (vla-get-FullName d))
    (if (and (eq (type fn) 'STR) (= (strcase fname) (strcase fn)))
      (setq ndoc d)
    )
  )
  (cond
    ((not ndoc)
     (list "失败" "打开图纸失败"))
    (t
     ;; 全程兜底：任何 COM 异常都转成"失败"消息，不中断批处理
     (setq res (vl-catch-all-apply 'tkxg:rename-doc (list ndoc pat newname fname)))
     ;; 无论成败都关闭文档（OK 时已存盘；失败时丢弃改动）
     (vl-catch-all-apply 'vla-Close (list ndoc :vlax-false))
     (vlax-release-object ndoc)
     (if (vl-catch-all-error-p res)
       (setq res (list "失败"
                       (strcat "处理出错: " (vl-catch-all-error-message res)))))
     res
    )
  )
)

;; 处理单个文件，返回 ("状态" "消息")；状态 = OK / 跳过 / 失败
;; 处理会话中已打开的图纸：直接在打开的文档上改
;; 干净图纸(无未保存修改) -> 改完自动存盘(先备份)；
;; 有未保存修改 -> 只改名不代存（*tkxg:nosave*），提示手动保存
(defun tkxg:process-ondoc (fname bakfile pat newname / docs ndoc fn res savedflag)
  (setq docs (vla-get-Documents (vlax-get-acad-object))
        ndoc nil)
  (vlax-for d docs
    (setq fn (vla-get-FullName d))
    (if (and (eq (type fn) 'STR) (= (strcase fname) (strcase fn)))
      (setq ndoc d)
    )
  )
  (if (not ndoc)
    (list "失败" "定位已打开文档失败")
    (progn
      ;; Saved=T 图纸干净；=nil 或取不到状态时按有修改处理（只改名不代存）
      (setq savedflag (vl-catch-all-apply 'vla-get-Saved (list ndoc))
            *tkxg:nosave* (not (and (not (vl-catch-all-error-p savedflag))
                                    (= :vlax-true savedflag))))
      (if (and (not *tkxg:nosave*)
               (not (vl-file-copy fname
                                  (tkxg:fixpath bakfile))))
        (progn
          (setq *tkxg:nosave* nil)
          (list "失败" (strcat "备份失败(目标:" (tkxg:fixpath bakfile) ")"))
        )
        (progn
          (setq res (vl-catch-all-apply 'tkxg:rename-doc
                      (list ndoc pat newname fname)))
          (setq *tkxg:nosave* nil)
          (if (vl-catch-all-error-p res)
            (list "失败" (strcat "处理出错: " (vl-catch-all-error-message res)))
            res
          )
        )
      )
    )
  )
)

(defun tkxg:process (fname bakfile pat newname)
  (cond
    ;; 会话内已打开的图纸：直接在打开文档上改（不再跳过）
    ((tkxg:isopen fname) (tkxg:process-ondoc fname bakfile pat newname))
    ((not (vl-file-systime fname)) (list "跳过" "文件不存在或不可读"))
  ;; 先备份（保留原扩展名），备份失败绝不改源文件
  ((not (vl-file-copy fname
                      (tkxg:fixpath bakfile)))
   (list "失败"
         (strcat "备份失败(源:" fname " 目标:" (tkxg:fixpath bakfile) ")")))
    (t (tkxg:process-open fname pat newname))
  )
)

;; ---------- 根文件夹四入口选择（TKGM/TKGPP 共用） --------------------
;; 回车=上次目录(注册表记忆) B=浏览对话框 D=当前图纸目录 F=选文件定位
;; 返回根目录全路径或 nil
(defun tkxg:getroot (/ k f root)
  (if (null *tkxg:root*) (setq *tkxg:root* (vl-registry-read *tkxg:reg* "RootDir")))
  (initget "Browse Drawing File")
  (setq k (getkword
            (strcat "\n根文件夹 [回车=上次目录 浏览(B) 当前图纸目录(D) 选文件定位(F)]"
                    (if (and *tkxg:root* (vl-file-directory-p *tkxg:root*))
                      (strcat " <" *tkxg:root* ">")
                      " <浏览>")
                    ": ")))
  (cond
    ((= k "Browse")
     (setq root (tkxg:browsefolder "选择要处理的根文件夹")))
    ((= k "Drawing")
     (setq root (getvar "DWGPREFIX"))
     ;; DWGPREFIX 自带尾部反斜杠，剥掉
     (while (and (> (strlen root) 1) (= (substr root (strlen root) 1) "\\"))
       (setq root (substr root 1 (1- (strlen root))))
     )
     (if (= root "") (setq root nil))
     (if (not root) (tkxg:say "当前图纸未保存，取不到所在目录。"))
    )
    ((= k "File")
     (setq f (getfiled "选择根文件夹内任意一张 DWG/DXF 图纸（可取消）"
                       (if *tkxg:dir* (strcat *tkxg:dir* "\\") "")
                       "dwg;dxf" 0))
     (if f
       (progn
         (if (and (not (vl-string-search "\\" f)) (not (vl-string-search "/" f)))
           (setq f (strcat (getvar "DWGPREFIX") f))
         )
         (setq root (tkxg:parentdir f))
         (if (= (substr root (strlen root) 1) ":")
           (setq root (strcat root "\\"))
         )
       )
     )
    )
    (t
     (setq root (if (and *tkxg:root* (vl-file-directory-p *tkxg:root*))
                  *tkxg:root*
                  nil))
     (if (not root)
       (progn
         (tkxg:say "上次目录不可用，改用浏览对话框。")
         (setq root (tkxg:browsefolder "选择要处理的根文件夹"))
       )
     )
    )
  )
  (if root
    (progn
      (setq *tkxg:root* root *tkxg:dir* root)
      (vl-registry-write *tkxg:reg* "RootDir" root)
    )
  )
  root
)

;; 存盘链：Save 失败时用带显式格式参数的 SaveAs（DWG/DXF 都要带类型，
;; ZWCAD 不带类型会弹对话框/报 Automation 错误）
;; 返回 nil=成功；否则返回错误消息字符串
(defun tkxg:savedoc (ndoc fname / serr sv2 saved tp)
  (setq serr (vl-catch-all-apply 'vla-Save (list ndoc)))
  (if (not (vl-catch-all-error-p serr))
    nil
    (progn
      (setq sv2 (vl-catch-all-error-message serr))
      (if (/= (strcase (vl-filename-extension fname)) ".DXF")
        (strcat "存盘出错: " sv2)
        (progn
          (setq saved nil)
          (foreach tp (quote (65 61 49 37 25 13 1))
            (if (not saved)
              (progn
                (setq serr (vl-catch-all-apply 'vla-SaveAs (list ndoc fname tp)))
                (if (not (vl-catch-all-error-p serr)) (setq saved t))
              )
            )
          )
          (if saved nil (strcat "存盘出错: " sv2))
        )
      )
    )
  )
)

;; 块参照: 若含 项目名称/项目名称1 属性对则修整（单行无对 -> nil 不动）
;; 返回: "ok" / "already" / "same" / nil
(defun tkxg:trypair (blkref scale / onm atts att m1 m2)
  (setq onm (vl-catch-all-apply 'vla-get-ObjectName (list blkref)))
  (if (or (vl-catch-all-error-p onm) (/= onm "AcDbBlockReference"))
    nil
    (progn
      (setq atts (vl-catch-all-apply 'vlax-invoke (list blkref 'GetAttributes)))
      (if (vl-catch-all-error-p atts) (setq atts nil))
      (setq m1 nil
            m2 nil)
      (foreach att atts
        (cond
          ((= (strcase (vla-get-TagString att)) (strcase "项目名称"))
           (setq m1 att))
          ((= (strcase (vla-get-TagString att)) (strcase "项目名称1"))
           (setq m2 att))
        )
      )
      (if (and m1 m2) (tkxg:fixpair m1 m2 scale) nil)
    )
  )
)

;; 就近找工作包 XML（root\工作包、上级\工作包、root 本层）；多个列编号选
(defun tkxg:pickxml (root / dirs dir lst f cands)
  (setq cands nil
        dirs  (list (strcat root "\\工作包")
                    (strcat (tkxg:parentdir root) "\\工作包")
                    root))
  (foreach dir dirs
    (if (vl-file-directory-p dir)
      (foreach f (vl-directory-files dir "*.xml" 1)
        (setq f (strcat dir "\\" f))
        (if (not (member (strcase f) (mapcar 'strcase cands)))
          (setq cands (cons f cands))
        )
      )
    )
  )
  (setq cands (reverse cands))
  (cond
    ((null cands) nil)
    ((null (cdr cands)) (car cands))
    (t
     (princ (strcat "\n  找到 " (itoa (length cands)) " 个工作包 XML，请选择："))
     (setq i 0)
     (foreach f cands
       (setq i (1+ i))
       (princ (strcat "\n    [" (itoa i) "] " f))
     )
     (setq i (getint "\n  输入序号(0=取消) <1>: "))
     (cond
       ((not i) (car cands))
       ((and (> i 0) (<= i (length cands))) (nth (1- i) cands))
       (t nil)
     )
    )
  )
)

;; 生成 EZ_SMEDI_OPENTASKWITHFILE 助手 VBS:
;; 等浏览窗出现 -> 工作包路径先写剪贴板再 Ctrl+V 粘贴（SendKeys 打不出中文）
;; -> 回车打开 -> 轮询"提示"框关闭 -> 写完成标记
(defun tkxg:writeopenvbs (vbs xmlpath marker / fh)
  (setq fh (open vbs "w"))
  (write-line "' TKGPP opentask helper (clipboard paste, Chinese-safe)" fh)
  (write-line "Option Explicit" fh)
  (write-line "On Error Resume Next" fh)   ; 任何一步出错也保证走完并写标记
  (write-line "Dim ws, fso, ts, tf, ps, i, lg" fh)
  (write-line "Set ws  = CreateObject(\"WScript.Shell\")" fh)
  (write-line "Set fso = CreateObject(\"Scripting.FileSystemObject\")" fh)
  (write-line "lg = ws.ExpandEnvironmentStrings(\"%TEMP%\") & \"\\tkgm_vbs.log\"" fh)
  (write-line "Sub L(s)" fh)
  (write-line "  Dim f" fh)
  (write-line "  Set f = fso.OpenTextFile(lg, 8, True)" fh)
  (write-line "  f.WriteLine Now & \"  \" & s" fh)
  (write-line "  f.Close" fh)
  (write-line "End Sub" fh)
  (write-line "L \"start\"" fh)
  (write-line "tf = ws.ExpandEnvironmentStrings(\"%TEMP%\") & \"\\tkgm_path.txt\"" fh)
  (write-line "Set ts = fso.CreateTextFile(tf, True)" fh)
  (write-line (strcat "ts.Write \"" xmlpath "\"") fh)
  (write-line "ts.Close" fh)
  (write-line "L \"path file written\"" fh)
  ;; 用 PowerShell Set-Clipboard 按 ANSI(GBK) 读文件（clip.exe 会按 UTF-8 误解 GBK）
  (write-line "ps = \"powershell -NoProfile -Command \" & Chr(34) & \"Set-Clipboard -Value (Get-Content -Raw -Encoding Default -LiteralPath '\" & tf & \"')\" & Chr(34)" fh)
  (write-line "ws.Run ps, 0, False" fh)     ; 非阻塞：后面等"打开"窗口的轮询足够剪贴板就绪
  (write-line "L \"clipboard set, err=\" & Err.Number" fh)
  ;; 自适应等待"打开"浏览窗（最长 60s，200ms 粒度）；出现后再填路径，不盲发
  (write-line "i = 0" fh)
  (write-line "Do While i < 300" fh)
  (write-line "  WScript.Sleep 200" fh)
  (write-line "  i = i + 1" fh)
  (write-line "  If ws.AppActivate(\"打开\") Then Exit Do" fh)
  (write-line "Loop" fh)
  (write-line "L \"open-dialog poll i=\" & i" fh)
  (write-line "If i < 300 Then" fh)
  (write-line "  WScript.Sleep 400" fh)     ; 留出剪贴板/对话框就绪
  (write-line "  ws.SendKeys \"%n\"" fh)    ; Alt+N 激活"文件名"输入框
  (write-line "  WScript.Sleep 150" fh)
  (write-line "  ws.SendKeys \"^a\"" fh)
  (write-line "  WScript.Sleep 80" fh)
  (write-line "  ws.SendKeys \"^v\"" fh)    ; 粘贴完整路径
  (write-line "  WScript.Sleep 200" fh)
  (write-line "  ws.SendKeys \"~\"" fh)     ; 回车打开
  (write-line "  L \"path pasted and entered\"" fh)
  (write-line "End If" fh)
  (write-line "L \"path pasted and entered\"" fh)
  (write-line "i = 0" fh)
  (write-line "Do While i < 120" fh)
  (write-line "  WScript.Sleep 1000" fh)
  (write-line "  i = i + 1" fh)
  (write-line "  If ws.AppActivate(\"提示\") Then" fh)
  (write-line "    WScript.Sleep 300" fh)
  (write-line "    ws.SendKeys \"~\"" fh)
  (write-line "    L \"tip box closed at \" & i" fh)
  (write-line "    Exit Do" fh)
  (write-line "  End If" fh)
  (write-line "Loop" fh)
  (write-line "L \"done, tip poll i=\" & i" fh)
  (write-line (strcat "fso.CreateTextFile \"" marker "\", True") fh)
  (write-line "L \"marker written\"" fh)
  (close fh)
  vbs
)

;; 调试日志（追加写 %TEMP%\tkgm_lisp.log，用于定位阶段性问题）
(defun tkxg:dbg (s / f)
  (setq f (open (strcat (tkxg:fixpath (getenv "TEMP")) "\\tkgm_lisp.log") "a"))
  (if f
    (progn
      (write-line (strcat (rtos (getvar "DATE") 2 8) "  " s) f)
      (close f)
    )
  )
)

;; 生成 EZ_SMEDI_ALTERFRAME 助手 VBS（工作包已打开时用；Tabx3+回车点"刷新图框信息"）
(defun tkxg:writealtervbs (vbs marker / fh)
  (setq fh (open vbs "w"))
  (write-line "' TKGPP alterframe helper" fh)
  (write-line "Option Explicit" fh)
  (write-line "On Error Resume Next" fh)
  (write-line "Dim ws, fso, i" fh)
  (write-line "Set ws = CreateObject(\"WScript.Shell\")" fh)
  (write-line "WScript.Sleep 2500" fh)
  (write-line "ws.SendKeys \"{TAB}{TAB}{TAB}\"" fh)
  (write-line "WScript.Sleep 400" fh)
  (write-line "ws.SendKeys \"~\"" fh)
  (write-line "WScript.Sleep 300" fh)
  (write-line "ws.SendKeys \"~\"" fh)
  (write-line "i = 0" fh)
  (write-line "Do While i < 120" fh)
  (write-line "  WScript.Sleep 1000" fh)
  (write-line "  i = i + 1" fh)
  (write-line "  If ws.AppActivate(\"提示\") Then" fh)
  (write-line "    WScript.Sleep 300" fh)
  (write-line "    ws.SendKeys \"~\"" fh)
  (write-line "    Exit Do" fh)
  (write-line "  End If" fh)
  (write-line "Loop" fh)
  (write-line "Set fso = CreateObject(\"Scripting.FileSystemObject\")" fh)
  (write-line (strcat "fso.CreateTextFile \"" marker "\", True") fh)
  (close fh)
  vbs
)

;; 轮询完成标记（用 DATE 做真实秒计时；自旋空转计时不准会提前超时）
(defun tkxg:waitmarker (marker secs / t0 ok el)
  (setq ok nil
        t0 (getvar "DATE")
        el 0.0)
  (while (and (not ok) (< el secs))
    (if (findfile marker) (setq ok t))
    (setq el (* (- (getvar "DATE") t0) 86400.0))
    (if (and (not ok) (< el secs))
      (repeat 20000 (setq el el))          ; 轻量让出，避免满 CPU
    )
  )
  ok
)

;; ---------- 目标图纸内执行：折行项目名称居中（TKGPP 经 SendCommand 调用）----
;; 用 ssget 在"目标图纸自身上下文"里扫全图块参照——绕开 ZWCAD 未注册的
;; vla-get-Blocks/Layouts 集合访问器；结果写全局 + 完成标记文件
(defun c:TKGPPFIX (/ ss i e obj cnt skp err res f)
  (setq cnt 0
        skp 0
        err 0)
  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e   (ssname ss i)
              i   (1+ i))
        (setq obj (vl-catch-all-apply 'vlax-ename->vla-object (list e)))
        (if (not (vl-catch-all-error-p obj))
          (progn
            (setq res (vl-catch-all-apply 'tkxg:trypair (list obj *tkxg:scale*)))
            (cond
              ((vl-catch-all-error-p res) (setq err (1+ err)))
              ((= res "ok") (setq cnt (1+ cnt)))
              ((or (= res "already") (= res "same")) (setq skp (1+ skp)))
            )
          )
        )
      )
    )
  )
  (setq *tkxg:pc*   cnt
        *tkxg:psk*  skp
        *tkxg:perr* err)
  (setq f (open (strcat (tkxg:fixpath (getenv "TEMP")) "\\tkgm_fix.flag") "w"))
  (if f
    (progn
      (write-line (strcat (itoa cnt) " " (itoa skp) " " (itoa err)) f)
      (close f)
    )
  )
  (princ)
)

;; ---------- 主命令 --------------------------------------------------

(defun c:TKGM (/ f root files rec pat newname bakroot stamp k full bakfile
                 rel log okc skipc failc fh logfile pn i n xmlf xf xcands)
  (vl-load-com)
  (princ "\n========== 图框块批量改名 TKGM v8.5 ==========")
  ;; 记忆上次输入（注册表持久化，仅匹配模式/工作包路径有默认值）
  (if (null *tkxg:pat*) (setq *tkxg:pat* (vl-registry-read *tkxg:reg* "Pattern")))
  (if (null *tkxg:xml*) (setq *tkxg:xml* (vl-registry-read *tkxg:reg* "Xml")))

  ;; 1. 选根文件夹（四入口，详见 tkxg:getroot）
  (setq root (tkxg:getroot))
  (if (not root)
    (tkxg:say "已取消。")
    (progn
      ;; 2. 是否包含子文件夹
      (initget "Yes No")
      (setq rec (getkword "\n是否处理所有子文件夹? [是(Y)/否(N)] <是>: "))
      (setq rec (/= rec "No"))                     ; 回车默认 = 含子文件夹
      ;; 3. 收集文件清单
      (setq files (tkxg:collect root rec))
      (if (not files)
        (tkxg:say (strcat "该文件夹内没有 DWG/DXF 文件: " root))
        (progn
          ;; 4. 块名匹配模式（回车用默认）+ 新块名（必须手动输入）
          (setq pat (getstring (strcat "\n要修改的图框块名(支持通配符 * ?) <"
                                       (if *tkxg:pat* *tkxg:pat* "*-S-A3-1")
                                       ">: ")))
          (if (= pat "") (setq pat (if *tkxg:pat* *tkxg:pat* "*-S-A3-1")))
          (setq *tkxg:pat* pat)
          (vl-registry-write *tkxg:reg* "Pattern" pat)
          ;; 新块名：直接输入；或回车=按工作包项目编号自动生成
          (setq *tkxg:pnum* nil)
          (setq newname (getstring T
                          "\n新图框块名(直接输入；回车=按工作包项目编号自动生成): "))
          (if (= newname "")
            (progn
              ;; 在当前文件夹附近找工作包 XML（多个则列出让用户选）
              (setq xcands (tkxg:findxml root))
              (cond
                ;; 没找到：弹文件框（默认定位到上次工作包所在目录）
                ((null xcands)
                 (setq xf (getfiled "选择工作包 XML 文件"
                            (if *tkxg:xml*
                              (strcat (tkxg:parentdir *tkxg:xml*) "\\")
                              (strcat root "\\"))
                            "xml" 0))
                 (if xf (setq xmlf (vl-string-translate "/" "\\" xf)))
                )
                ;; 只有一个：直接用
                ((null (cdr xcands))
                 (setq xmlf (car xcands))
                 (tkxg:say (strcat "工作包: " xmlf))
                )
                ;; 多个：列编号让用户选，回车默认第 1 个
                (t
                 (princ (strcat "\n  找到 " (itoa (length xcands))
                                " 个工作包 XML，请选择："))
                 (setq i 0)
                 (foreach xf xcands
                   (setq i (1+ i))
                   (princ (strcat "\n    [" (itoa i) "] " xf))
                 )
                 (setq pick (getint "\n  输入序号(0=取消) <1>: "))
                 (cond
                   ((not pick) (setq xmlf (car xcands)))
                   ((and (> pick 0) (<= pick (length xcands)))
                    (setq xmlf (nth (1- pick) xcands)))
                 )
                )
              )
              (if xmlf
                (progn
                  (setq *tkxg:pnum* (tkxg:getprojno xmlf))
                  (if *tkxg:pnum*
                    (progn
                      (setq *tkxg:xml* xmlf)
                      (vl-registry-write *tkxg:reg* "Xml" xmlf)
                      (tkxg:say (strcat "项目编号: " *tkxg:pnum*
                                        "  （" xmlf "）"))
                    )
                    (tkxg:say "未能从工作包解出项目编号。")
                  )
                )
              )
              ;; 自动生成失败 -> 回退为必须手动输入
              (while (and (not *tkxg:pnum*) (= newname ""))
                (setq newname (getstring T "\n新图框块名(必须输入): "))
                (if (= newname "") (tkxg:say "新块名不能为空，请重新输入。"))
              )
            )
          )
          ;; 5. 显示执行计划并确认
          (princ (strcat "\n  ----------------------------------------"
                         "\n  根文件夹   : " root
                         "\n  处理范围   : " (if rec "含所有子文件夹" "仅本层文件夹")
                         "\n  图纸数量   : " (itoa (length files)) " 张"
                         "\n  块名匹配   : " pat
                         "\n  新块名     : " (if *tkxg:pnum*
                                              (strcat "自动生成（项目编号 " *tkxg:pnum* " + 原尾段）")
                                              newname)
                         "\n  ----------------------------------------"))
          (initget "Yes No")
          (setq k (getkword
                    "\n将逐张打开-改名-存盘-关闭（自动备份，期间图纸窗口会闪动），确认执行? [是(Y)/否(N)] <是>: "))
          (if (/= k "No")                            ; 回车默认 = 是
            (progn
              ;; 6. 建备份根目录（镜像相对路径）
              (setq stamp   (tkxg:stamp)
                    bakroot (strcat root "\\备份_" stamp))
              (tkxg:mkdirs bakroot)
              (tkxg:say (strcat "备份目录: " bakroot))
              (setq *tkxg:scrlog* logfile)
              ;; 屏蔽代理对象提示，批完恢复
              (setq pn (getvar "PROXYNOTICE"))
              (setvar "PROXYNOTICE" 0)
              ;; 7. 逐张处理
              (setq okc 0 skipc 0 failc 0 log '()
                    n (length files) i 0)
              (foreach rel files
                (setq i       (1+ i)
                      full    (tkxg:fixpath (strcat root "\\" rel))
                      bakfile (tkxg:fixpath (strcat bakroot "\\" rel)))
                (tkxg:mkdirs (tkxg:parentdir bakfile))
                (princ (strcat "\n  [" (itoa i) "/" (itoa n) "] " rel))
                (setq res  (tkxg:process full bakfile pat newname)
                      log  (cons (strcat (car res) "\t" rel "\t" (cadr res)) log))
                (cond
                  ((= (car res) "OK")   (setq okc   (1+ okc)))
                  ((= (car res) "跳过") (setq skipc (1+ skipc)))
                  (t                    (setq failc (1+ failc))))
                (princ (strcat " -> [" (car res) "] " (cadr res)))
              )
              (setvar "PROXYNOTICE" pn)
              (setq log (reverse log))
              ;; 8. 写修改日志
              (setq logfile (strcat root "\\图框名修改日志_" stamp ".txt")
                    fh      (open logfile "w"))
              (write-line "图框块批量改名 TKGM 修改日志" fh)
              (write-line (strcat "根文件夹: " root
                                  "    范围: " (if rec "含子文件夹" "仅本层")) fh)
              (write-line (strcat "块名匹配: " pat "    新块名: "
                                  (if *tkxg:pnum*
                                    (strcat "自动(项目编号 " *tkxg:pnum* ")")
                                    newname)) fh)
              (write-line "----------------------------------------" fh)
              (foreach l log (write-line l fh))
              (write-line "----------------------------------------" fh)
              (write-line (strcat "成功 " (itoa okc) " / 跳过 " (itoa skipc)
                                  " / 失败 " (itoa failc)) fh)
              (close fh)
              (princ (strcat "\n  ----------------------------------------"
                             "\n  完成: 成功 " (itoa okc)
                             "  跳过 " (itoa skipc)
                             "  失败 " (itoa failc)
                             "\n  修改日志: " logfile))
            )
            (tkxg:say "已取消，未做任何修改。")
          )
        )
      )
    )
  )
  (princ)
)

;; ---------- 图框文字修整：折行项目名称两行居中（TKGMJZ） --------------
;; 场景: EZ_SMEDI_ALTERFRAME 刷新后，过长项目名称折成两行且未竖向居中
;; 对象: 图框块内 TAG=项目名称/项目名称1 的单行属性对（中中对齐）
;; 规则: 第一行中心上移 2.5x比例，第二行中心 = 第一行原中心 - 2.5x比例
;;       （两行间距 = 5x比例，绕原位置对称居中）；间距已达标则跳过

;; 修整一对单行属性（TAG: 项目名称 / 项目名称1，中中对齐）
;; 返回: "ok" / "already"(间距已是目标) / "same"(两行同Y，不动) / nil(出错)
(defun tkxg:fixpair (a1 a2 scale / p1 p2 gap y1 dy2)
  (setq p1 (vlax-get a1 'TextAlignmentPoint)
        p2 (vlax-get a2 'TextAlignmentPoint))
  ;; 对齐点为 (0,0,0) 时退回插入点
  (if (and (= (car p1) 0.0) (= (cadr p1) 0.0) (= (caddr p1) 0.0))
    (setq p1 (vlax-get a1 'InsertionPoint)))
  (if (and (= (car p2) 0.0) (= (cadr p2) 0.0) (= (caddr p2) 0.0))
    (setq p2 (vlax-get a2 'InsertionPoint)))
  (setq gap (- (cadr p1) (cadr p2)))
  (cond
    ;; 两行 Y 一致：不动（老板要求）
    ((< (abs gap) 0.01) "same")
    ;; 间距已达标：幂等跳过
    ((< (abs (- gap (* 5.0 scale))) 0.01) "already")
    (t
     (setq y1  (cadr p1)
           dy2 (- (- y1 (* 2.5 scale)) (cadr p2)))
     (vla-put-TextAlignmentPoint
       a1 (vlax-3d-point (car p1) (+ (cadr p1) (* 2.5 scale)) (caddr p1)))
     (vla-put-TextAlignmentPoint
       a2 (vlax-3d-point (car p2) (+ (cadr p2) dy2) (caddr p2)))
     "ok"
    )
  )
)

(defun c:TKGMJZ (/ ss scale sv i n obj atts att cnt skip fail m1 m2 res)
  (vl-load-com)
  (princ "\n========== 图框文字修整 TKGMJZ ==========")
  (if (null *tkxg:scale*)
    (progn
      ;; ZWCAD 读不到值时返回 nil，atof(nil) 会炸，必须判型
      (setq sv (vl-registry-read *tkxg:reg* "Scale"))
      (setq *tkxg:scale* (if (eq (type sv) 'STR) (atof sv) 1.0))
    )
  )
  (setq scale (getreal (strcat "\n图框比例(距离按此缩放) <"
                               (rtos *tkxg:scale* 2 0) ">: ")))
  (if (not scale) (setq scale *tkxg:scale*))
  (setq *tkxg:scale* scale)
  (vl-registry-write *tkxg:reg* "Scale" (rtos scale 2 0))
  (princ "\n  选择图框块(可框选全图，自动识别含属性对的图框): ")
  (setq ss (ssget))
  (if (not ss)
    (tkxg:say "未选择。")
    (progn
      (setq i 0 n (sslength ss) cnt 0 fail 0 skip 0)
      (while (< i n)
        (setq obj (vlax-ename->vla-object (ssname ss i))
              i   (1+ i))
        (cond
          ;; 块参照: 找 项目名称/项目名称1 属性对修整
          ((= (vla-get-ObjectName obj) "AcDbBlockReference")
           (setq atts (vlax-invoke obj 'GetAttributes)
                 m1   nil
                 m2   nil)
           (foreach att atts
             (cond
               ((= (strcase (vla-get-TagString att)) (strcase "项目名称"))
                (setq m1 att))
               ((= (strcase (vla-get-TagString att)) (strcase "项目名称1"))
                (setq m2 att))
             )
           )
           (if (and m1 m2)
             (progn
              (setq res (tkxg:fixpair m1 m2 scale))
              (cond
                ((= res "ok")      (setq cnt (1+ cnt)))
                ((= res "already") (setq skip (1+ skip)))
                ((= res "same")    (setq skip (1+ skip)))
                (t                 (setq fail (1+ fail)))
              )
             )
           )
          )
        )
      )
      (princ (strcat "\n  完成: 调整 " (itoa cnt) " 处"
                     (if (> skip 0) (strcat "，已是目标间距 " (itoa skip) " 处") "")
                     (if (> fail 0) (strcat "，失败 " (itoa fail) " 处") "")
                     "\n  提示: 若刷新图框后文字重新生成，需再跑一次 TKGMJZ"))
    )
  )
  (princ)
)

;; ---------- 跨文件批量：改名 + 刷新图框 + 折行居中（TKGPP） ----------
;; 每张图: 打开 -> ①改名(可工作包自动) -> ②EZ_SMEDI_OPENTASKWITHFILE
;;         (VBS 自动把工作包 XML 路径填进浏览窗; 队列预置回车应答"全选")
;;         -> ③项目名称对折行居中 -> ④存盘关闭
;; 期间会自动移动鼠标键盘焦点，请勿操作电脑
;; 纯 LISP：枚举块表，返回匹配 pat 的块名列表（排除匿名块与外部参照）
(defun tkxg:matchblocks (pat / nm out bn fl)
  (setq out nil
        nm  (tblnext "BLOCK" T))
  (while nm
    (setq bn (cdr (assoc 2 nm))
          fl (cdr (assoc 70 nm)))
    (if (and (eq (type bn) 'STR)
             (/= (substr bn 1 1) "*")
             (= 0 (logand (if fl fl 0) 13))       ; 1=匿名 4=外部参照 8=覆盖
             (wcmatch (strcase bn) (strcase pat)))
      (setq out (cons bn out))
    )
    (setq nm (tblnext "BLOCK"))
  )
  (reverse out)
)

;; 当前图纸内：项目名称/项目名称1 属性对折行居中（ssget 全图，不碰 COM 集合）
;; 返回 (调整数 跳过数 错误数)
(defun tkxg:fixcurrent (scale / ss i e obj cnt skp err res)
  (setq cnt 0
        skp 0
        err 0)
  (setq ss (ssget "_X" '((0 . "INSERT"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i)
              i (1+ i))
        (setq obj (vl-catch-all-apply 'vlax-ename->vla-object (list e)))
        (if (not (vl-catch-all-error-p obj))
          (progn
            (setq res (vl-catch-all-apply 'tkxg:trypair (list obj scale)))
            (cond
              ((vl-catch-all-error-p res) (setq err (1+ err)))
              ((= res "ok") (setq cnt (1+ cnt)))
              ((or (= res "already") (= res "same")) (setq skp (1+ skp)))
            )
          )
        )
      )
    )
  )
  (list cnt skp err)
)

;; ---------- 脚本模式辅助（TKGPP 经 .scr 逐张调用） --------------------
;; 追加写批处理日志
(defun tkxg:logapp (s / f)
  (setq f (open *tkxg:scrlog* "a"))
  (if f
    (progn
      (write-line s f)
      (close f)
    )
  )
)

;; 读日志行（汇总用）
(defun tkxg:readlog (p / f out l)
  (setq out nil)
  (if (setq f (open p "r"))
    (progn
      (while (setq l (read-line f))
        (setq out (cons l out))
      )
      (close f)
    )
  )
  (reverse out)
)

;; 脚本内：改名（当前图纸，纯 LISP 块表 + -RENAME）
(defun tkxg:scr-ren (rel / bl nn)
  (if (and *tkxg:pat* (/= *tkxg:pat* ""))
    (progn
      (setq bl (tkxg:matchblocks *tkxg:pat*))
      (cond
        ((null bl)
         (tkxg:logapp (strcat "跳过\t" rel "\t改名：无匹配块")))
        ((cdr bl)
         (tkxg:logapp (strcat "跳过\t" rel "\t改名：多个块名 "
                              (tkxg:join bl ", "))))
        (t
         (setq nn (if *tkxg:pnum*
                    (tkxg:genname (car bl) *tkxg:pnum*)
                    *tkxg:newname*))
         (if nn
           (progn
             (command "_.-RENAME" "_B" (car bl) nn)
             (tkxg:logapp (strcat "OK\t" rel "\t改名 " (car bl) " -> " nn)))
           (tkxg:logapp (strcat "跳过\t" rel "\t改名：无法生成新名"))
         )
        )
      )
    )
    (tkxg:logapp (strcat "跳过\t" rel "\t跳过改名"))
  )
  (princ)
)

;; 脚本内：派发刷新助手 VBS（配合下一条 EZ_SMEDI_OPENTASKWITHFILE）
(defun tkxg:scr-res ()
  (if (findfile *tkxg:doneflag*) (vl-file-delete *tkxg:doneflag*))
  (startapp "wscript.exe" *tkxg:openvbs*)
  (princ)
)

;; 脚本内：折行居中（当前图纸，ssget 全图）
(defun tkxg:scr-cen (/ r)
  (setq r (tkxg:fixcurrent *tkxg:scale*))
  (tkxg:logapp (strcat "OK\t居中 调整 " (itoa (nth 0 r)) " 跳过 " (itoa (nth 1 r))))
  (princ)
)

;; 脚本内：关闭当前图纸（COM 关闭，脚本可继续下一张）
(defun tkxg:scr-end ()
  (tkxg:logapp "OK\t已存盘关闭")
  (vl-catch-all-apply 'vla-Close
    (list (vla-get-ActiveDocument (vlax-get-acad-object)) :vlax-false))
  (princ)
)

;; ---------- 刷新通道（COM 打开 + 命令 + VBS，已验证可用；不进脚本）---------
(defun tkxg:refreshpass (root files / docs ndoc fn full i n marker firstflag)
  (setq marker (strcat (tkxg:fixpath (getenv "TEMP")) "\\tkgm_done.flag"))
  (setq i 0
        n (length files))
  (foreach rel files
    (setq i    (1+ i)
          full (tkxg:fixpath (strcat root "\\" rel)))
    (princ (strcat "\n  刷新 [" (itoa i) "/" (itoa n) "] " rel))
    (setq docs (vla-get-Documents (vlax-get-acad-object))
          ndoc nil)
    (vl-catch-all-apply 'vla-Open (list docs full))
    (vlax-for d docs
      (setq fn (vla-get-FullName d))
      (if (and (eq (type fn) 'STR) (= (strcase full) (strcase fn)))
        (setq ndoc d)
      )
    )
    (if ndoc
      (progn
        (if (findfile marker) (vl-file-delete marker))
        ;; 每张图都用 OPENTASKWITHFILE（每张图都要加载工作包+刷新，均有浏览窗）
        (startapp "wscript.exe" *tkxg:openvbs*)
        (command "EZ_SMEDI_OPENTASKWITHFILE")
        (if (tkxg:waitmarker marker 180)
          (tkxg:logapp (strcat "OK\t" rel "\t刷新完成"))
          (tkxg:logapp (strcat "失败\t" rel "\t刷新超时(检查工作包)"))
        )
        (command "_.QSAVE")
        (vl-catch-all-apply 'vla-Close (list ndoc :vlax-false))
      )
      (tkxg:logapp (strcat "失败\t" rel "\t打开失败"))
    )
  )
  (princ)
)


;; ---------- 对象级 COM：对指定文档做 项目名称 折行居中 ------------------
;; 返回 (调整数 跳过数 错误数)；无法枚举块表时返回 (-1 0 0)
(defun tkxg:center-doc (ndoc scale / blox blk nm e res cnt skp err)
  (setq cnt 0
        skp 0
        err 0)
  (setq blox (vl-catch-all-apply 'vla-get-Blocks (list ndoc)))
  (if (vl-catch-all-error-p blox)
    (list -1 0 0)
    (progn
      (vlax-for blk blox
        (setq nm (vl-catch-all-apply 'vla-get-Name (list blk)))
        (if (and (eq (type nm) 'STR)
                 (or (= nm "*Model_Space")
                     (vl-string-search "*Paper_Space" nm)))
          (vlax-for e blk
            (setq res (vl-catch-all-apply 'tkxg:trypair (list e scale)))
            (cond
              ((vl-catch-all-error-p res) (setq err (1+ err)))
              ((= res "ok") (setq cnt (1+ cnt)))
              ((or (= res "already") (= res "same")) (setq skp (1+ skp)))
            )
          )
        )
      )
      (list cnt skp err)
    )
  )
)
(defun c:TKGPP (/ root files scale pat newname bakroot stamp k full bakfile
                  rel i n fh logfile sv xcands xmlf oxml docs ndoc fn
                  res1 r pc psk perr setr serr)
  (vl-load-com)
  (princ "\n========== 图框批量处理(改名+居中) TKGPP v8.5 ==========")
  (princ "\n  （刷新请先在 iMIS 里完成；本命令做 改名+居中+存盘+关闭）")
  (if (null *tkxg:pat*) (setq *tkxg:pat* (vl-registry-read *tkxg:reg* "Pattern")))
  (if (null *tkxg:xml*) (setq *tkxg:xml* (vl-registry-read *tkxg:reg* "Xml")))

  ;; 1. 根文件夹（四入口）
  (if (not (setq root (tkxg:getroot)))
    (tkxg:say "已取消。")
    (progn
      ;; 2. 收集文件（含子文件夹）
      (setq files (vl-remove-if
                    '(lambda (s) (= (substr s 1 2) "~$"))
                    (tkxg:collect root t)))
      (if (not files)
        (tkxg:say (strcat "该文件夹内没有 DWG/DXF 文件: " root))
        (progn
          ;; 3. 比例
          (if (null *tkxg:scale*)
            (progn
              (setq sv (vl-registry-read *tkxg:reg* "Scale"))
              (setq *tkxg:scale* (if (eq (type sv) 'STR) (atof sv) 1.0))
            )
          )
          (setq scale (getreal (strcat "\n图框比例(居中距离按此缩放) <"
                                       (rtos *tkxg:scale* 2 0) ">: ")))
          (if (not scale) (setq scale *tkxg:scale*))
          (setq *tkxg:scale* scale)
          (vl-registry-write *tkxg:reg* "Scale" (rtos scale 2 0))
          ;; 4. 改名设置（留空 = 跳过改名）
          (setq pat (getstring (strcat "\n要修改的图框块名(支持通配符 * ?；回车=跳过改名) <"
                                       (if *tkxg:pat* *tkxg:pat* "*-S-A3-1") ">: ")))
          (if (/= pat "")
            (progn
              (setq *tkxg:pat* pat)
              (vl-registry-write *tkxg:reg* "Pattern" pat)
              (setq *tkxg:pnum* nil)
              (setq newname (getstring T
                              "\n新图框块名(直接输入；回车=按工作包项目编号自动生成): "))
              (if (= newname "")
                (progn
                  (setq oxml (tkxg:pickxml root))
                  (if oxml
                    (progn
                      (setq *tkxg:pnum* (tkxg:getprojno oxml))
                      (if *tkxg:pnum*
                        (progn
                          (setq *tkxg:xml* oxml)
                          (vl-registry-write *tkxg:reg* "Xml" oxml)
                          (tkxg:say (strcat "项目编号: " *tkxg:pnum*)))
                        (tkxg:say "未能从工作包解出项目编号。")))
                    (tkxg:say "未找到工作包 XML。"))
                  (while (and (not *tkxg:pnum*) (= newname ""))
                    (setq newname (getstring T "\n新图框块名(必须输入): "))
                    (if (= newname "") (tkxg:say "新块名不能为空，请重新输入。"))
                  ))))
            (setq *tkxg:pnum* nil
                  newname    "")
          )
          ;; 5. 确认
          (princ (strcat "\n  ----------------------------------------"
                         "\n  根文件夹   : " root
                         "\n  图纸数量   : " (itoa (length files)) " 张（含子文件夹）"
                         "\n  改名       : " (if (/= pat "")
                                              (strcat pat " -> "
                                                (if *tkxg:pnum*
                                                  (strcat "自动(项目编号 " *tkxg:pnum* " + 原尾段)")
                                                  newname))
                                              "跳过")
                         "\n  折行居中   : 项目名称对 上移2.5/间距5 x比例"
                         (rtos scale 2 0)
                         "\n  ----------------------------------------"))
          (initget "Yes No")
          (setq k (getkword "\n确认执行? [是(Y)/否(N)] <是>: "))
          (if (/= k "No")
            (progn
              ;; 6. 备份 + 日志
              (setq stamp   (tkxg:stamp)
                    bakroot (tkxg:fixpath (strcat root "\\备份_" stamp))
                    logfile (strcat root "\\图框批量处理日志_" stamp ".txt"))
              (tkxg:mkdirs bakroot)
              (tkxg:say (strcat "备份目录: " bakroot))
              (setq fh (open logfile "w"))
              (write-line "图框批量处理 TKGPP v8.5（改名+居中+存盘+关闭）" fh)
              (write-line (strcat "根文件夹: " root "    比例: " (rtos scale 2 0)) fh)
              (write-line (strcat "块名匹配: " pat "    新块名: "
                          (if *tkxg:pnum*
                            (strcat "自动(项目编号 " *tkxg:pnum* ")")
                            newname)) fh)
              (write-line "----------------------------------------" fh)
              (close fh)
              ;; 7. 逐张处理（全部对象级 COM，不用 (command ...)）
              (setq i 0
                    n (length files))
              (foreach rel files
                (setq i       (1+ i)
                      full    (tkxg:fixpath (strcat root "\\" rel))
                      bakfile (tkxg:fixpath (strcat bakroot "\\" rel)))
                (princ (strcat "\n  [" (itoa i) "/" (itoa n) "] " rel))
                (tkxg:mkdirs (tkxg:parentdir bakfile))
                (cond
                  ((not (vl-file-systime full))
                   (princ " -> [跳过] 文件不存在")
                   (tkxg:logapp (strcat "跳过\t" rel "\t文件不存在")))
                  ((not (vl-file-copy full bakfile))
                   (princ " -> [失败] 备份失败")
                   (tkxg:logapp (strcat "失败\t" rel "\t备份失败")))
                  (t
                   ;; 打开（对象级 COM）
                   (setq docs (vla-get-Documents (vlax-get-acad-object))
                         ndoc nil)
                   (vl-catch-all-apply 'vla-Open (list docs full))
                   (vlax-for d docs
                     (setq fn (vla-get-FullName d))
                     (if (and (eq (type fn) 'STR) (= (strcase full) (strcase fn)))
                       (setq ndoc d)
                     )
                   )
                   (if (not ndoc)
                     (progn
                       (princ " -> [失败] 打开失败")
                       (tkxg:logapp (strcat "失败\t" rel "\t打开失败")))
                     (progn
                       ;; ① 改名（对象级 COM）
                       (setq res1 (if (/= pat "")
                                    (vl-catch-all-apply 'tkxg:rename-doc
                                      (list ndoc pat newname full))
                                    (list "跳过" "跳过改名")))
                       (if (vl-catch-all-error-p res1)
                         (setq res1 (list "失败" (vl-catch-all-error-message res1))))
                       (princ (strcat "\n    ①改名: [" (car res1) "] " (cadr res1)))
                       (tkxg:logapp (strcat "改名" (car res1) "\t" rel "\t" (cadr res1)))
                       ;; ② 居中（对象级 COM）
                       (setq r (tkxg:center-doc ndoc scale))
                       (setq pc (nth 0 r) psk (nth 1 r) perr (nth 2 r))
                       (if (= pc -1)
                         (progn
                           (princ "\n    ②居中: [失败] 无法枚举块表(COM)")
                           (tkxg:logapp (strcat "失败\t" rel "\t居中:无法枚举块表")))
                         (progn
                           (princ (strcat "\n    ②居中: 调整 " (itoa pc)
                                          " 处，跳过 " (itoa psk) " 处"))
                           (tkxg:logapp (strcat "OK\t" rel "\t居中 调整 "
                                               (itoa pc) " 跳过 " (itoa psk))))
                       )
                       ;; ③ 存盘（对象级 COM，跳过无改动）
                       (setq setr (vl-catch-all-apply 'vla-get-Saved (list ndoc)))
                       (if (and (not (vl-catch-all-error-p setr))
                                (= :vlax-true setr))
                         (progn
                           (princ "\n    ③存盘: 跳过（无改动）")
                           (tkxg:logapp (strcat "OK\t" rel "\t无改动")))
                         (progn
                           (setq serr (tkxg:savedoc ndoc full))
                           (if serr
                             (progn
                               (princ (strcat "\n    ③存盘: [失败] " serr))
                               (tkxg:logapp (strcat "失败\t" rel "\t" serr)))
                             (progn
                               (princ "\n    ③存盘: OK")
                               (tkxg:logapp (strcat "OK\t" rel "\t存盘")))
                           )
                         )
                       )
                       ;; ④ 关闭（对象级 COM）
                       (vl-catch-all-apply 'vla-Close (list ndoc :vlax-false))
                       (vlax-release-object ndoc)
                     )
                   )
                  )
                )
              )
              (princ (strcat "\n  ----------------------------------------"
                             "\n  完成，日志: " logfile))
            )
            (tkxg:say "已取消。")
          )
        )
      )
    )
  )
  (princ)
)





;; ---------- 当前图纸一键：改名(可选) + 折行居中 + 存盘（方案B，最稳）----
;; 用法：在打开的图纸里执行 TKGMP1；配合 iMIS 的刷新流程使用
(defun c:TKGMP1 (/ pat scale newname xf bl nn r)
  (vl-load-com)
  (princ "\n========== 当前图纸处理 TKGMP1 ==========")
  (if (null *tkxg:scale*) (setq *tkxg:scale* 1.0))
  (setq scale (getreal (strcat "\n图框比例 <" (rtos *tkxg:scale* 2 0) ">: ")))
  (if (not scale) (setq scale *tkxg:scale*))
  (setq *tkxg:scale* scale)
  (setq pat (getstring (strcat "\n要改名的图框块(通配符；回车=不改名) <"
                               (if *tkxg:pat* *tkxg:pat* "*-S-A3-1") ">: ")))
  (if (/= pat "")
    (progn
      (setq *tkxg:pat* pat)
      (setq newname (getstring T "\n新图框块名(回车=选工作包 XML 自动生成): "))
      (setq *tkxg:pnum* nil)
      (if (= newname "")
        (progn
          (setq xf (getfiled "选择工作包 XML（取消则手动输入新名）"
                             (getvar "DWGPREFIX") "xml" 0))
          (if xf
            (progn
              (setq *tkxg:pnum* (tkxg:getprojno (vl-string-translate "/" "\\" xf)))
              (if *tkxg:pnum* (princ (strcat "\n项目编号: " *tkxg:pnum*)))
            )
          )
          (while (and (not *tkxg:pnum*) (= newname ""))
            (setq newname (getstring T "\n新图框块名(必须输入): "))
            (if (= newname "") (princ "\n新块名不能为空。"))
          )
        )
      )
      (setq bl (tkxg:matchblocks pat)
            nn (if *tkxg:pnum*
                 (if (car bl) (tkxg:genname (car bl) *tkxg:pnum*) nil)
                 newname))
      (cond
        ((null bl) (princ "\n[跳过] 无匹配的图框块"))
        ((cdr bl)  (princ (strcat "\n[跳过] 匹配到多个块名: " (tkxg:join bl ", "))))
        ((not nn)  (princ "\n[跳过] 无法生成新名"))
        ((= (strcase (car bl)) (strcase nn))
                   (princ (strcat "\n[跳过] " (car bl) " 已是新名")))
        ((tblsearch "BLOCK" nn)
                   (princ (strcat "\n[跳过] 目标块名已存在(" nn ")")))
        (t         (command "_.-RENAME" "_B" (car bl) nn)
                   (princ (strcat "\n改名: " (car bl) " -> " nn)))
      )
    )
    (princ "\n[跳过] 不改名")
  )
  (setq r (tkxg:fixcurrent scale))
  (princ (strcat "\n居中: 调整 " (itoa (nth 0 r)) " 处，跳过 " (itoa (nth 1 r)) " 处"))
  (command "_.QSAVE")
  (princ "\n已存盘(QSAVE)。")
  (princ)
)
;; ---------- 供外部脚本(VBS)调用：改名 / 居中（参数走注册表 HKCU\Software\TKGM）----------
(defun tkxg:regget (n d / r)
  (setq r (vl-catch-all-apply 'vl-registry-read (list *tkxg:reg* n)))
  (if (and (not (vl-catch-all-error-p r)) (eq (type r) 'STR)) r d)
)

;; 无匹配块时的排查提示：列出图内所有含 "S-A" 的块名（最多 8 个）
(defun tkxg:candhint (/ nm bn fl lst)
  (setq lst nil
        nm  (tblnext "BLOCK" T))
  (while nm
    (setq bn (cdr (assoc 2 nm))
          fl (cdr (assoc 70 nm)))
    (if (and (eq (type bn) 'STR)
             (/= (substr bn 1 1) "*")
             (= 0 (logand (if fl fl 0) 13))   ; 1=匿名 4=外部参照 8=覆盖
             (wcmatch (strcase bn) "*S-A*"))
      (setq lst (cons bn lst))
    )
    (setq nm (tblnext "BLOCK"))
  )
  (setq lst (reverse lst))
  (if lst
    (strcat " 候选(" (itoa (length lst)) "):" (tkxg:join lst ","))
    " 候选:无(图内没有含 S-A 的块)")
)

;; 图面上该块名的插入实例数（0 = 只有块定义、图面没用，多半是历史遗留垃圾块）
;; 注意：ssget "_X" 只搜当前空间，会漏掉布局里的图框；故优先用 ActiveX 遍历
;; 模型空间 + 所有布局，失败再回退 ssget。
(defun tkxg:countinserts (bn / cnt r ss)
  (setq cnt 0)
  (vl-load-com)
  (setq r (vl-catch-all-apply
            '(lambda ()
               (vlax-for lo (vla-get-Layouts (vla-get-ActiveDocument (vlax-get-acad-object)))
                 (vlax-for o (vla-get-Block lo)
                   (if (and (= (vla-get-ObjectName o) "AcDbBlockReference")
                            (= (strcase (vla-get-Name o)) (strcase bn)))
                     (setq cnt (1+ cnt))
                   )
                 )
               )
               cnt
             )
          )
  )
  (if (vl-catch-all-error-p r)
    (progn
      (setq ss (ssget "_X" (list '(0 . "INSERT") (cons 2 bn))))
      (if ss (sslength ss) 0)
    )
    (if (numberp r) r 0)
  )
)

;; (块名 . 实例数) 列表 -> "块A(3),块B(0)"
(defun tkxg:joindet (lst / s d)
  (setq s "")
  (foreach d lst
    (setq s (strcat s (if (= s "") "" ",") (car d) "(" (itoa (cdr d)) ")"))
  )
  s
)

;; 按「优先关键字」从块名清单里筛：返回名字包含关键字(不分大小写)的子集
;; 关键字为空时原样返回（等于不筛）
(defun tkxg:filterpref (lst pref / r s)
  (setq r nil)
  (if (and pref (/= pref ""))
    (progn
      (setq s (strcase pref))
      (foreach x lst (if (wcmatch (strcase x) (strcat "*" s "*")) (setq r (cons x r))))
      (reverse r)
    )
    lst
  )
)

;; 本张图改名（外部调用）；结果写 %TEMP%\tkgm_ren.flag
;; flag 格式：状态|说明
;;   OK / ALREADY / EXIST / NONAME / NOPAT
;;   NOMATCH      没有一个块匹配模式
;;   MULTI_USED   多个块匹配且图面都用到 → 无法判定，跳过（除非用 Prefer 关键字定出唯一一个）
;;   ALL_UNUSED   多个块匹配但图面都没用（全是遗留块定义）→ 跳过
(defun c:TKGREN (/ pat nn xml pn bl f msg old det d used prf sel)
  (vl-load-com)
  (setq pat (tkxg:regget "Pattern" "")
        nn  (tkxg:regget "NewName" "")
        xml (tkxg:regget "Xml" "")
        prf (tkxg:regget "Prefer" "")
        pn  nil
        msg "")
  (if (and (= nn "") xml (findfile xml))
    (setq pn (tkxg:getprojno xml))
  )
  (cond
    ((= pat "")
     (setq msg "NOPAT|未填改名模式")
     (princ "\nTKGREN: 跳过改名(模式为空)"))
    (t
     (setq bl (tkxg:matchblocks pat))
     (cond
       ((null bl)
        (setq msg (strcat "NOMATCH|无块匹配 " pat (tkxg:candhint)))
        (princ "\nTKGREN: 无匹配块"))
       ((cdr bl)
        ;; 多个块匹配时，按「图面实际插入使用」筛选：真正要改的是被用到的那个，
        ;; 其余多为历次复制图纸残留的块定义垃圾。
        (setq det nil)
        (foreach d bl (setq det (cons (cons d (tkxg:countinserts d)) det)))
        (setq det (reverse det))
        (setq used nil)
        (foreach d det (if (> (cdr d) 0) (setq used (cons (car d) used))))
        (setq used (reverse used))
        ;; 图面用到多个时，若填了「优先关键字」(Prefer)，按关键字在「有用」的块里再筛一次
        (setq old nil sel nil)
        (cond
          ((null used)
           (setq msg (strcat "ALL_UNUSED|" (tkxg:joindet det)))
           (princ "\nTKGREN: 多个匹配块但图面均未使用，跳过"))
          ((null (cdr used))
           (setq old (car used) sel ""))
          (t
           (setq sel (tkxg:filterpref used prf))
           (cond
             ((= prf "")
              (setq msg (strcat "MULTI_USED|" (tkxg:joindet det)))
              (princ "\nTKGREN: 多个块名且图面都用，无法判定，跳过"))
             ((null sel)
              (setq msg (strcat "MULTI_USED|" (tkxg:joindet det)
                                "  [优先关键字[" prf "]无命中]"))
              (princ "\nTKGREN: 多个块名且优先关键字无命中，跳过"))
             ((cdr sel)
              (setq msg (strcat "MULTI_USED|" (tkxg:joindet det)
                                "  [优先关键字[" prf "]命中" (itoa (length sel)) "个]"))
              (princ "\nTKGREN: 优先关键字命中多个，跳过"))
             (t
              (setq old (car sel))
              (princ (strcat "\nTKGREN: 按优先关键字[" prf "]选中 " old)))
           )
          )
        )
        (if old
          (progn
           (setq nn (if pn (tkxg:genname old pn) nn))
           (cond
             ((or (not nn) (= nn ""))
              (setq msg (strcat "NONAME|匹配到 " old " 但生成不出新名"))
              (princ "\nTKGREN: 无新名，已跳过"))
             ((= (strcase old) (strcase nn))
              (setq msg (strcat "ALREADY|" nn
                                (if (and prf (/= prf "")) (strcat "  [按优先关键字[" prf "]选中]") "")))
              (princ (strcat "\nTKGREN: " old " 已是新名，跳过")))
             ((tblsearch "BLOCK" nn)
              (setq msg (strcat "EXIST|" nn " 已存在，用到的是 " old))
              (princ (strcat "\nTKGREN: 目标块名已存在(" nn ")，跳过")))
             (t
              (command "_.-RENAME" "_B" old nn)
              (setq msg (strcat "OK|" old " -> " nn
                                "  [原匹配" (itoa (length bl)) "块" 
                                (if (and prf (/= prf "")) (strcat "，按优先关键字[" prf "]选中" ) "，选图面在用者")
                                "]"))
              (princ (strcat "\nTKGREN: " old " -> " nn)))
           )
          )
        )
        )
       (t
        (setq old (car bl)
              nn  (if pn (tkxg:genname old pn) nn))
        (cond
          ((or (not nn) (= nn ""))
           (setq msg (strcat "NONAME|匹配到 " old " 但生成不出新名")
                 nn  nil)
           (princ "\nTKGREN: 无新名，已跳过"))
          ((= (strcase old) (strcase nn))
           (setq msg (strcat "ALREADY|" nn))
           (princ (strcat "\nTKGREN: " old " 已是新名，跳过")))
          ((tblsearch "BLOCK" nn)
           (setq msg (strcat "EXIST|" nn " 已存在，原块 " old))
           (princ (strcat "\nTKGREN: 目标块名已存在(" nn ")，跳过")))
          (t
           (command "_.-RENAME" "_B" old nn)
           (setq msg (strcat "OK|" old " -> " nn))
           (princ (strcat "\nTKGREN: " old " -> " nn)))
        )
       )
     )
    )
  )
  (setq f (open (strcat (tkxg:fixpath (getenv "TEMP")) "\\tkgm_ren.flag") "w"))
  (if f
    (progn
      (write-line (if (= msg "") "NONE|无结果" msg) f)
      (close f)
    )
  )
  (princ)
)

;; 清理本张图里「匹配改名模式但图面没有实例」的图框块定义（历史遗留垃圾块）
;; 安全边界：只删 0 实例的块；PURGE 本身也会拒绝被嵌套引用的块，图面引用一律不动。
;; 结果写 %TEMP%\tkgm_purge.flag
(defun c:TKGPURGE (/ pat bl bn n f done lst)
  (vl-load-com)
  (setq pat (tkxg:regget "Pattern" "")
        n 0 done nil lst nil)
  (if (= pat "")
    (setq done "NOPAT|未填模式，跳过清理")
    (progn
      (setq bl (tkxg:matchblocks pat))
      (foreach bn bl
        (if (= 0 (tkxg:countinserts bn))
          (progn
            (command "_.-PURGE" "_B" bn "_N")
            (setq n (1+ n))
            (setq lst (cons bn lst))
          )
        )
      )
      (setq lst (reverse lst))
      (setq done (strcat "清 " (itoa n) " 个" (if lst (strcat ":" (tkxg:join lst ",")) "")))
    )
  )
  (setq f (open (strcat (tkxg:fixpath (getenv "TEMP")) "\\tkgm_purge.flag") "w"))
  (if f (progn (write-line done f) (close f)))
  (princ (strcat "\nTKGPURGE: " done))
  (princ)
)

;; 本张图折行居中（外部调用）；结果写 %TEMP%\tkgm_fix.flag
(defun c:TKGFIX (/ sc r f)
  (setq sc (atof (tkxg:regget "Scale" "1")))
  (if (<= sc 0.0) (setq sc 1.0))
  (setq r (tkxg:fixcurrent sc))
  (princ (strcat "\nTKGFIX: 调整 " (itoa (nth 0 r)) " 处，跳过 " (itoa (nth 1 r)) " 处"))
  (setq f (open (strcat (tkxg:fixpath (getenv "TEMP")) "\\tkgm_fix.flag") "w"))
  (if f
    (progn
      (write-line (strcat (itoa (nth 0 r)) " " (itoa (nth 1 r))) f)
      (close f)
    )
  )
  (princ)
)

(princ "\n图框块批量改名 v8.5 已加载：TKGM 批量改名 | TKGMJZ 折行居中 | TKGPP 跨文件批量 | TKGMP1 当前图纸一键 | TKGPR 批量刷新 | TKGDIAG 自检 | TKGREN/TKGFIX 外部调用入口。")
(princ);; ---------- 批量刷新图框（逐张 打开→刷新→存盘→关闭）TKGPR ----------
;; 刷新走 EZ_SMEDI_OPENTASKWITHFILE + VBS 填工作包路径；
;; 每张图先校验"已成为活动图纸"再发命令，避免命令/QSAVE 落到别的图上
;; 等待指定路径成为活动图纸（vla-Open 后 ActiveDocument 有延迟）；返回 t/nil
(defun tkxg:waitactive (full secs / t0 el act ok)
  (setq ok nil
        t0 (getvar "DATE")
        el 0.0)
  (while (and (not ok) (< el secs))
    (setq act (vl-catch-all-apply 'vla-get-ActiveDocument
                                   (list (vlax-get-acad-object))))
    (if (not (vl-catch-all-error-p act))
      (progn
        (setq act (vl-catch-all-apply 'vla-get-FullName (list act)))
        (if (and (eq (type act) 'STR)
                 (= (strcase full) (strcase act)))
          (setq ok t)
        )
      )
    )
    (if (not ok)
      (progn
        (repeat 400000 (setq el el))     ; 轻量让出，降低轮询频率
        (setq el (* (- (getvar "DATE") t0) 86400.0))
      )
    )
  )
  ok
)



(defun c:TKGPR (/ root files marker tmpp openvbs docs ndoc fn actn setr serr setok opres dcnt setok
                  full rel bakroot stamp logfile i n bakdir)
  (vl-load-com)
  (princ "\n========== 批量刷新图框 TKGPR v8.5 ==========")
  (if (null *tkxg:xml*) (setq *tkxg:xml* (vl-registry-read *tkxg:reg* "Xml")))
  (if (not (setq root (tkxg:getroot)))
    (tkxg:say "已取消。")
    (progn
      (setq files (vl-remove-if
                    '(lambda (s) (= (substr s 1 2) "~$"))
                    (tkxg:collect root t)))
      (if (not files)
        (tkxg:say (strcat "该文件夹内没有 DWG/DXF 文件: " root))
        (progn
          (setq tmpp    (tkxg:fixpath (getenv "TEMP"))
                marker  (strcat tmpp "\\tkgm_done.flag")
                openvbs (strcat tmpp "\\tkgm_opentask.vbs"))
          (if (not (and *tkxg:xml* (findfile *tkxg:xml*)))
            (setq *tkxg:xml* (tkxg:pickxml root))
          )
          (if (not *tkxg:xml*)
            (tkxg:say "未找到工作包 XML，已取消。请把工作包 XML 放到根文件夹后重跑。")
            (progn
              (tkxg:writeopenvbs openvbs *tkxg:xml* marker)
              (princ (strcat "\n  ----------------------------------------"
                             "\n  根文件夹 : " root
                             "\n  图纸数量 : " (itoa (length files)) " 张（含子文件夹）"
                             "\n  工作包   : " *tkxg:xml*
                             "\n  动作     : 逐张 打开→刷新图框→存盘→关闭"
                             "\n  ----------------------------------------"
                             "\n  ★ 期间会自动操作对话框，请勿移动鼠标键盘！★"))
              (initget "Yes No")
              (if (not (= (getkword "\n确认执行? [是(Y)/否(N)] <是>: ") "No"))
                (progn
                  (setq stamp   (tkxg:stamp)
                        bakroot (tkxg:fixpath (strcat root "\\备份_" stamp))
                        logfile (strcat root "\\批量刷新日志_" stamp ".txt")
                        *tkxg:scrlog* logfile)
                  (tkxg:mkdirs bakroot)
                  (setq i 0
                        n (length files))
                  (foreach rel files
                    (setq i    (1+ i)
                          full (tkxg:fixpath (strcat root "\\" rel)))
                    (princ (strcat "\n  [" (itoa i) "/" (itoa n) "] " rel))
                    (setq bakdir (tkxg:parentdir (tkxg:fixpath (strcat bakroot "\\" rel))))
                    (tkxg:mkdirs bakdir)
                    (if (not (vl-file-copy full (tkxg:fixpath (strcat bakroot "\\" rel))))
                      (progn
                        (princ " -> [失败] 备份失败")
                        (tkxg:logapp (strcat "失败\t" rel "\t备份失败"))
                      )
                      (progn
                        ;; 打开（对象级 COM）
                        (setq docs (vla-get-Documents (vlax-get-acad-object))
                              ndoc nil)
                        (setq opres (vl-catch-all-apply 'vla-Open (list docs full)))
                        (setq dcnt (vl-catch-all-apply 'vla-get-Count (list docs)))
                        (tkxg:logapp (strcat "调试	" rel "	Open->"
                                             (if (vl-catch-all-error-p opres)
                                               (vl-catch-all-error-message opres)
                                               "ok")
                                             " 文档数->"
                                             (if (vl-catch-all-error-p dcnt)
                                               "?"
                                               (itoa dcnt))))
                        (vlax-for d docs
                          (setq fn (vla-get-FullName d))
                          (if (and (eq (type fn) 'STR) (= (strcase full) (strcase fn)))
                            (setq ndoc d)
                          )
                        )
                        ;; 把刚打开的文档切为活动图（Document.Activate），再轮询确认
                        (if ndoc (vl-catch-all-apply (quote vla-Activate) (list ndoc)))
                        ;; 等待该图成为活动图纸（vla-Open 后有延迟，轮询最多20s）
                        (setq setok (and ndoc (tkxg:waitactive full 20)))
                        (if (not setok)
                          (progn
                            (setq actn "?")
                            (vl-catch-all-apply
                              (quote (lambda ()
                                 (setq actn (vla-get-FullName
                                              (vla-get-ActiveDocument (vlax-get-acad-object)))))))
                            (princ (strcat " -> [跳过] 等待活动图纸超时；当前活动图: "
                                           (if (eq (type actn) (quote STR)) actn "(取不到)")))
                            (tkxg:logapp (strcat "跳过	" rel "	活动图纸超时；当前: "
                                                 (if (eq (type actn) (quote STR)) actn "(取不到)")))
                            (if ndoc (vl-catch-all-apply (quote vla-Close) (list ndoc :vlax-false)))
                          )
                          (progn
                            ;; 刷新（VBS 填工作包 + 命令；此刻目标图是活动图）
                            (if (findfile marker) (vl-file-delete marker))
                            (startapp "wscript.exe" openvbs)
                            (command "EZ_SMEDI_OPENTASKWITHFILE")
                            (if (tkxg:waitmarker marker 180)
                              (progn
                                (princ "\n    刷新: 完成")
                                (tkxg:logapp (strcat "OK\t" rel "\t刷新完成"))
                              )
                              (progn
                                (princ "\n    刷新: [失败] 自动化超时(180s)")
                                (tkxg:logapp (strcat "失败\t" rel "\t刷新超时"))
                              )
                            )
                            ;; 存盘（COM；无改动则跳过）
                            (setq setr (vl-catch-all-apply 'vla-get-Saved (list ndoc)))
                            (if (and (not (vl-catch-all-error-p setr))
                                     (= :vlax-true setr))
                              (progn
                                (princ "\n    存盘: 跳过（无改动）")
                                (tkxg:logapp (strcat "OK\t" rel "\t无改动"))
                              )
                              (progn
                                (setq serr (tkxg:savedoc ndoc full))
                                (if serr
                                  (progn
                                    (princ (strcat "\n    存盘: [失败] " serr))
                                    (tkxg:logapp (strcat "失败\t" rel "\t" serr))
                                  )
                                  (progn
                                    (princ "\n    存盘: OK")
                                    (tkxg:logapp (strcat "OK\t" rel "\t存盘"))
                                  )
                                )
                              )
                            )
                            ;; 关闭（对象级 COM）
                            (vl-catch-all-apply 'vla-Close (list ndoc :vlax-false))
                            (vlax-release-object ndoc)
                            (princ "\n    已关闭，继续下一张")
                          )
                        )
                      )
                    )
                  )
                  (princ (strcat "\n  ----------------------------------------"
                                 "\n  完成，日志: " logfile))
                )
                (tkxg:say "已取消。")
              )
            )
          )
        )
      )
    )
  )
  (princ)
)




;; ---------- COM 自检 v2：Open 后能否把新文档切为活动图（TKGDIAG）----------
(defun c:TKGDIAG (/ acad docs f r n1 n2 act nm doc res d)
  (vl-load-com)
  (princ "\n========== CAD COM 自检 v2 TKGDIAG ==========")
  (setq acad (vlax-get-acad-object))
  (setq docs (vl-catch-all-apply 'vlax-get (list acad 'Documents)))
  (cond
    ((vl-catch-all-error-p docs)
     (princ (strcat "\n[1] 取 Documents 失败: " (vl-catch-all-error-message docs)))
    )
    (t
     (setq n1 (vl-catch-all-apply 'vla-get-Count (list docs)))
     (princ (strcat "\n[1] 当前文档数: "
                    (if (vl-catch-all-error-p n1) "(取不到)" (itoa n1))))
     (setq act (vl-catch-all-apply 'vla-get-ActiveDocument (list acad)))
     (if (not (vl-catch-all-error-p act))
       (progn
         (setq nm (vl-catch-all-apply 'vla-get-FullName (list act)))
         (princ (strcat "\n[2] 活动图: " (if (eq (type nm) 'STR) nm "(取不到)")))
       )
     )
     (setq f (getfiled "选一张 DWG 做打开测试（可取消）" "" "dwg" 0))
     (cond
       ((not f)
        (princ "\n[3] 未选文件，测试结束"))
       (t
        (setq f (vl-string-translate "/" "\\" f))
        (setq r (vl-catch-all-apply 'vla-Open (list docs f)))
        (princ (strcat "\n[3] Open 返回: "
                       (if (vl-catch-all-error-p r)
                         (vl-catch-all-error-message r)
                         "OK")))
        (setq n2 (vl-catch-all-apply 'vla-get-Count (list docs)))
        (princ (strcat "\n[4] 打开后文档数: "
                       (if (vl-catch-all-error-p n2) "(取不到)" (itoa n2))))
        ;; 找到刚打开的那个文档对象
        (setq doc nil)
        (vlax-for d docs
          (setq nm (vl-catch-all-apply 'vla-get-FullName (list d)))
          (if (and (eq (type nm) 'STR)
                   (= (strcase nm) (strcase f)))
            (setq doc d)
          )
        )
        (princ (strcat "\n[5] 找到目标文档对象: " (if doc "是" "否")))
        ;; 尝试一：Document.Activate
        (if doc
          (progn
            (setq res (vl-catch-all-apply 'vla-Activate (list doc)))
            (princ (strcat "\n[6] vla-Activate -> "
                           (if (vl-catch-all-error-p res)
                             (vl-catch-all-error-message res)
                             "OK")))
            (setq act (vl-catch-all-apply 'vla-get-ActiveDocument (list acad)))
            (if (not (vl-catch-all-error-p act))
              (progn
                (setq nm (vl-catch-all-apply 'vla-get-FullName (list act)))
                (princ (strcat "\n[7] 之后的活动图: "
                               (if (eq (type nm) 'STR) nm "(取不到)")))
              )
            )
          )
        )
        ;; 尝试二：Application.ActiveDocument 赋值
        (if doc
          (progn
            (setq res (vl-catch-all-apply 'vla-put-ActiveDocument (list acad doc)))
            (princ (strcat "\n[8] put-ActiveDocument -> "
                           (if (vl-catch-all-error-p res)
                             (vl-catch-all-error-message res)
                             "OK")))
            (setq act (vl-catch-all-apply 'vla-get-ActiveDocument (list acad)))
            (if (not (vl-catch-all-error-p act))
              (progn
                (setq nm (vl-catch-all-apply 'vla-get-FullName (list act)))
                (princ (strcat "\n[9] 之后的活动图: "
                               (if (eq (type nm) 'STR) nm "(取不到)")))
              )
            )
          )
        )
       )
     )
    )
  )
  (princ "\n=========================================")
  (princ)
)

