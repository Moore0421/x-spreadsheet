import { h } from './element';
import { cssPrefix } from '../config';
import { bindClickoutside, unbindClickoutside } from './event';

export default class TreeEditor {
  constructor(options = {}) {
    this.options = options;
    this.width = options.width || 500;
    this.height = options.height || 450;
    
    // 当前编辑的单元格信息
    this.currentCell = null;
    this.currentRi = null;
    this.currentCi = null;
    
    // 修改主容器样式
    this.el = h('div', `${cssPrefix}-tree-editor`)
      .css('width', `${this.width}px`)
      .css('height', `${this.height}px`)
      .css('position', 'fixed') // 改为 fixed 定位
      .css('background', '#fff')
      .css('box-shadow', '0 2px 12px rgba(0,0,0,0.15)')
      .css('z-index', '10000')
      .css('border-radius', '4px')
      .css('top', '50%')
      .css('left', '50%')
      .css('transform', 'translate(-50%, -50%)')
      .hide();
    
    // 创建标题栏
    this.header = h('div', `${cssPrefix}-tree-editor-header`)
      .css('padding', '10px')
      .css('border-bottom', '1px solid #e0e0e0')
      .css('display', 'flex')
      .css('justify-content', 'space-between')
      .css('align-items', 'center')
      .css('cursor', 'move');
    
    this.title = h('div', `${cssPrefix}-tree-editor-title`)
      .css('font-weight', 'bold')
      .html('编辑树形结构');
    
    this.closeBtn = h('div', `${cssPrefix}-tree-editor-close`)
      .css('cursor', 'pointer')
      .css('font-size', '16px')
      .html('×')
      .on('click', () => this.hide());
    
    this.header.children(this.title, this.closeBtn);
    
    // 创建内容区域
    this.content = h('div', `${cssPrefix}-tree-editor-content`)
      .css('padding', '10px')
      .css('height', 'calc(100% - 110px)')
      .css('overflow', 'auto');
    
    // 新增：提示词输入框
    this.placeholderInput = h('input', `${cssPrefix}-tree-editor-placeholder`)
      .css('width', '100%')
      .css('margin-bottom', '8px')
      .css('padding', '6px 8px')
      .css('box-sizing', 'border-box')
      .css('border', '1px solid #ccc')
      .css('border-radius', '4px')
      .attr('placeholder', '请输入提示词（如"点击选择"）');
    
    // 创建树形数据输入区域
    this.treeDataTextarea = h('textarea', `${cssPrefix}-tree-editor-textarea`)
      .css('width', '100%')
      .css('height', '88%')
      .css('resize', 'none')
      .css('padding', '8px')
      .css('box-sizing', 'border-box')
      .css('border', '1px solid #ccc')
      .css('border-radius', '4px')
      .attr('placeholder', '请输入JSON格式的树形数据，例如：\n[\n  {\n    "label": "一级选项1",\n    "value": "option1",\n    "children": [\n      {\n        "label": "二级选项1",\n        "value": "option1-1"\n      }\n    ]\n  }\n]');
    
    // 把提示词输入框插入到内容区最前面
    this.content.children(this.placeholderInput, this.treeDataTextarea);
    
    // 创建底部按钮区域
    this.footer = h('div', `${cssPrefix}-tree-editor-footer`)
      .css('padding', '10px')
      .css('border-top', '1px solid #e0e0e0')
      .css('display', 'flex')
      .css('justify-content', 'flex-end');
    
    this.cancelBtn = h('button', `${cssPrefix}-tree-editor-button ${cssPrefix}-tree-editor-cancel`)
      .css('padding', '5px 15px')
      .css('margin-right', '10px')
      .css('border', '1px solid #ddd')
      .css('background', '#f5f5f5')
      .css('border-radius', '4px')
      .css('cursor', 'pointer')
      .html('取消')
      .on('click', () => this.hide());
    
    this.saveBtn = h('button', `${cssPrefix}-tree-editor-button ${cssPrefix}-tree-editor-save`)
      .css('padding', '5px 15px')
      .css('border', 'none')
      .css('background', '#4b89ff')
      .css('color', '#fff')
      .css('border-radius', '4px')
      .css('cursor', 'pointer')
      .html('保存')
      .on('click', () => this.save());
    
    this.footer.children(this.cancelBtn, this.saveBtn);
    
    // 组装所有部分
    this.el.children(this.header, this.content, this.footer);
    
    // 回调函数
    this.change = () => {};

    // 添加拖拽功能
    this.initDrag();
  }
  
  // 设置树形数据
  setTreeData(treeData) {
    try {
      const jsonString = JSON.stringify(treeData || [], null, 2);
      this.treeDataTextarea.val(jsonString);
    } catch (e) {
      console.error('无法序列化树形数据:', e);
      this.treeDataTextarea.val('[]');
    }
    return this;
  }
  
  // 获取树形数据
  getTreeData() {
    try {
      return JSON.parse(this.treeDataTextarea.val() || '[]');
    } catch (e) {
      console.error('解析JSON失败:', e);
      return [];
    }
  }
  
  // 保存树形数据
  save() {
    if (!this.currentCell || this.currentRi === null || this.currentCi === null) {
      console.error('没有当前编辑的单元格信息');
      return;
    }
    
    const treeData = this.getTreeData();
    const placeholder = this.placeholderInput.val() || '';
    
    // 调用回调函数保存数据
    if (typeof this.change === 'function') {
      this.change({
        ri: this.currentRi,
        ci: this.currentCi,
        cell: this.currentCell,
        treeData: treeData,
        placeholder: placeholder,
      });
    } else {
      console.error('未设置回调函数');
    }
    
    this.hide();
  }
  
  // 初始化拖拽功能
  initDrag() {
    let isDragging = false;
    let startX, startY, initialX, initialY;

    this.header.on('mousedown', (e) => {
      if (e.target === this.closeBtn.el) return;
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      const rect = this.el.el.getBoundingClientRect();
      initialX = rect.left;
      initialY = rect.top;
      
      const mousemove = (e) => {
        if (!isDragging) return;
        e.preventDefault();
        
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;
        
        let newX = initialX + deltaX;
        let newY = initialY + deltaY;
        
        // 确保弹窗不会超出视口
        const maxX = window.innerWidth - this.el.el.offsetWidth;
        const maxY = window.innerHeight - this.el.el.offsetHeight;
        
        newX = Math.max(0, Math.min(newX, maxX));
        newY = Math.max(0, Math.min(newY, maxY));
        
        this.el
          .css('transform', 'none')
          .css('left', `${newX}px`)
          .css('top', `${newY}px`);
      };
      
      const mouseup = () => {
        isDragging = false;
        document.removeEventListener('mousemove', mousemove);
        document.removeEventListener('mouseup', mouseup);
      };
      
      document.addEventListener('mousemove', mousemove);
      document.addEventListener('mouseup', mouseup);
    });
  }

  // 修改设置位置的方法
  setOffset(offset) {
    if (offset) {
      // 确保弹窗显示在视口内
      const maxX = window.innerWidth - this.el.el.offsetWidth;
      const maxY = window.innerHeight - this.el.el.offsetHeight;
      
      const x = Math.max(0, Math.min(offset.left, maxX));
      const y = Math.max(0, Math.min(offset.top, maxY));
      
      this.el
        .css('transform', 'none')
        .css('left', `${x}px`)
        .css('top', `${y}px`);
    }
    return this;
  }
  
  // 显示编辑器
  show() {
    this.el.show();
    bindClickoutside(this.el);
    return this;
  }
  
  // 隐藏编辑器
  hide() {
    this.el.hide();
    unbindClickoutside(this.el);
    return this;
  }

  // 添加 setCell 方法到 TreeEditor 类
  setCell(ri, ci, cell) {
    this.currentRi = ri;
    this.currentCi = ci;
    this.currentCell = cell;
    
    // 设置树形数据
    if (cell && cell.treeData) {
      this.setTreeData(cell.treeData);
    } else {
      this.setTreeData([]);
    }
    
    // 设置提示词
    if (cell && typeof cell.text === 'string') {
      this.placeholderInput.val(cell.text);
    } else {
      this.placeholderInput.val('');
    }
    
    return this;
  }
}