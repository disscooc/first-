# 临时文件：新的合并函数，将被插入到 word_writer.py
# 用法：由 apply_patch.py 读取并替换

def _merge_yunda_rtype_col(tbl, comp_items, n_comp):
    """资源类型列（col 0）按 CPU/GPU 同类项垂直合并（同时支持单价/数量模式）。"""
    from docx.oxml.ns import qn
    from lxml import etree

    # 按 rtype 分组连续行
    groups = []
    current_group = []
    for idx in range(n_comp):
        rtype = comp_items[idx].rtype if idx < len(comp_items) else None
        if not current_group or current_group[-1][1] == rtype:
            current_group.append((idx, rtype))
        else:
            groups.append(current_group)
            current_group = [(idx, rtype)]
    if current_group:
        groups.append(current_group)

    for group in groups:
        if len(group) <= 1:
            continue  # 单行无需合并
        for gi, (row_idx, _rtype) in enumerate(group):
            row = tbl.rows[1 + row_idx]
            tcs = _get_real_tcs(row)
            if not tcs:
                continue
            tc = tcs[0]  # 资源类型 = col 0（两种模式都是 col 0）
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is None:
                tcPr = etree.SubElement(tc, qn('w:tcPr'))
                tc.insert(0, tcPr)
            # 清除旧的 vMerge
            old_vm = tcPr.find(qn('w:vMerge'))
            if old_vm is not None:
                tcPr.remove(old_vm)
            vm = etree.SubElement(tcPr, qn('w:vMerge'))
            vm.set(qn('w:val'), 'restart' if gi == 0 else 'continue')


def _merge_yunda_storage_col(tbl, n_comp, is_total):
    """存储资源列跨所有数据行垂直合并，不留空白格。
    is_total=True（单价模式）：列已删除重组，存储列 = tcs[3]
    is_total=False（数量模式）：7列完整，存储列 = tcs[5]
    """
    from docx.oxml.ns import qn
    from lxml import etree

    for idx in range(n_comp):
        row = tbl.rows[1 + idx]
        tcs = _get_real_tcs(row)
        if not tcs:
            continue
        # 根据模式选择列索引
        if is_total:
            col_idx = 3  # 单价模式：删除了 col 3,4,6，原 col 5 变为 col 3
        else:
            col_idx = 5  # 数量模式：7列完整，存储 = col 5
        if len(tcs) <= col_idx:
            continue
        tc = tcs[col_idx]
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn('w:tcPr'))
            tc.insert(0, tcPr)
        # 清除旧的 vMerge
        old_vm = tcPr.find(qn('w:vMerge'))
        if old_vm is not None:
            tcPr.remove(old_vm)
        vm = etree.SubElement(tcPr, qn('w:vMerge'))
        vm.set(qn('w:val'), 'restart' if idx == 0 else 'continue')


def _merge_yunda_total_col(tbl, n_comp):
    """数量模式下：总金额列（col 6）跨所有数据行垂直合并，只在第一行显示。"""
    from docx.oxml.ns import qn
    from lxml import etree

    for idx in range(n_comp):
        row = tbl.rows[1 + idx]
        tcs = _get_real_tcs(row)
        if not tcs or len(tcs) <= 6:
            continue
        tc = tcs[6]  # 总金额 = col 6（仅数量模式存在）
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = etree.SubElement(tc, qn('w:tcPr'))
            tc.insert(0, tcPr)
        # 清除旧的 vMerge
        old_vm = tcPr.find(qn('w:vMerge'))
        if old_vm is not None:
            tcPr.remove(old_vm)
        vm = etree.SubElement(tcPr, qn('w:vMerge'))
        vm.set(qn('w:val'), 'restart' if idx == 0 else 'continue')
