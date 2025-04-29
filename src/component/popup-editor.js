import { h } from './element';
import { cssPrefix } from '../config';
import { bindClickoutside, unbindClickoutside } from './event';

export default class PopupEditor {
  constructor(options = {}) {
    this.options = options;
    this.width = options.width || 800;
    this.height = options.height || 500;
    
    // 当前编辑的单元格信息
    this.currentCell = null;
    this.currentRi = null;
    this.currentCi = null;
    
    // 创建主容器
    this.el = h('div', `${cssPrefix}-popup-editor`)
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
    this.header = h('div', `${cssPrefix}-popup-editor-header`)
      .css('padding', '10px')
      .css('border-bottom', '1px solid #e0e0e0')
      .css('display', 'flex')
      .css('justify-content', 'space-between')
      .css('align-items', 'center')
      .css('cursor', 'move');
    
    this.title = h('div', `${cssPrefix}-popup-editor-title`)
      .css('font-weight', 'bold')
      .html('编辑弹窗内容');
    
    this.closeBtn = h('div', `${cssPrefix}-popup-editor-close`)
      .css('cursor', 'pointer')
      .css('font-size', '16px')
      .html('×')
      .on('click', () => this.hide());
    
    this.header.children(this.title, this.closeBtn);
    
    // 创建设置区域
    this.settingsArea = h('div', `${cssPrefix}-popup-editor-settings`)
      .css('padding', '10px')
      .css('border-bottom', '1px solid #e0e0e0')
      .css('display', 'flex')
      .css('gap', '10px');
    
    // 标题输入框
    this.titleLabel = h('label', `${cssPrefix}-popup-editor-label`)
      .css('display', 'flex')
      .css('align-items', 'center')
      .css('margin-right', '10px')
      .html('弹窗标题:');
    
    this.titleInput = h('input', `${cssPrefix}-popup-editor-title-input`)
      .css('padding', '5px')
      .css('border', '1px solid #ccc')
      .css('border-radius', '4px')
      .css('width', '200px')
      .attr('type', 'text')
      .attr('placeholder', '请输入弹窗标题');
    
    // 尺寸设置
    this.widthLabel = h('label', `${cssPrefix}-popup-editor-label`)
      .css('display', 'flex')
      .css('align-items', 'center')
      .css('margin-right', '5px')
      .css('margin-left', '10px')
      .html('宽度:');
    
    this.widthInput = h('input', `${cssPrefix}-popup-editor-width-input`)
      .css('padding', '5px')
      .css('border', '1px solid #ccc')
      .css('border-radius', '4px')
      .css('width', '60px')
      .attr('type', 'number')
      .attr('min', '200')
      .attr('max', '800')
      .attr('placeholder', '宽度');
    
    this.heightLabel = h('label', `${cssPrefix}-popup-editor-label`)
      .css('display', 'flex')
      .css('align-items', 'center')
      .css('margin-right', '5px')
      .css('margin-left', '10px')
      .html('高度:');
    
    this.heightInput = h('input', `${cssPrefix}-popup-editor-height-input`)
      .css('padding', '5px')
      .css('border', '1px solid #ccc')
      .css('border-radius', '4px')
      .css('width', '60px')
      .attr('type', 'number')
      .attr('min', '200')
      .attr('max', '600')
      .attr('placeholder', '高度');
    
    this.settingsArea.children(
      this.titleLabel, this.titleInput,
      this.widthLabel, this.widthInput,
      this.heightLabel, this.heightInput
    );
    
    // 创建内容区域
    this.content = h('div', `${cssPrefix}-popup-editor-content`)
      .css('padding', '10px')
      .css('height', 'calc(100% - 160px)') 
      .css('display', 'flex')
      .css('flex-direction', 'column')
      .css('overflow', 'auto'); // 添加滚动条支持
    
    this.contentLabel = h('label', `${cssPrefix}-popup-editor-content-label`)
      .css('margin-bottom', '5px')
      .css('font-weight', 'bold')
      .html('弹窗内容 (支持HTML):');
    
    // 创建HTML内容输入区域
    this.htmlTextarea = h('textarea', `${cssPrefix}-popup-editor-textarea`)
      .css('width', '100%')
      .css('flex', '1')
      .css('resize', 'none')
      .css('padding', '8px')
      .css('box-sizing', 'border-box')
      .css('border', '1px solid #ccc')
      .css('border-radius', '4px')
      .css('font-family', 'monospace')
      .attr('placeholder', '请输入HTML内容，例如：\n<div style="padding:10px">\n  <h3>标题</h3>\n  <p>这是一个<b>HTML</b>示例内容</p>\n  <div style="color:blue">这是蓝色文本</div>\n</div>');
    
    this.content.children(this.contentLabel, this.htmlTextarea);
    
    // 创建底部按钮区域
    this.footer = h('div', `${cssPrefix}-popup-editor-footer`)
      .css('padding', '10px')
      .css('border-top', '1px solid #e0e0e0')
      .css('display', 'flex')
      .css('justify-content', 'flex-end');
    
    this.cancelBtn = h('button', `${cssPrefix}-popup-editor-button ${cssPrefix}-popup-editor-cancel`)
      .css('padding', '5px 15px')
      .css('margin-right', '10px')
      .css('border', '1px solid #ddd')
      .css('background', '#f5f5f5')
      .css('border-radius', '4px')
      .css('cursor', 'pointer')
      .html('取消')
      .on('click', () => this.hide());
    
    this.saveBtn = h('button', `${cssPrefix}-popup-editor-button ${cssPrefix}-popup-editor-save`)
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
    this.el.children(this.header, this.settingsArea, this.content, this.footer);
    
    // 回调函数
    this.change = () => {};
    
    // 添加拖拽功能
    this.initDrag();
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

  // 设置弹窗数据
  setPopupData(popupData) {
    if (!popupData) {
      popupData = {
        title: '弹窗标题',
        width: 400,
        height: 300,
        content: '<div style="padding:10px">弹窗内容</div>'
      };
    }
    
    this.titleInput.val(popupData.title || '');
    this.widthInput.val(popupData.width || 400);
    this.heightInput.val(popupData.height || 300);
    this.htmlTextarea.val(popupData.content || '');
    
    return this;
  }
  
  // 获取弹窗数据
  getPopupData() {
    return {
      title: this.titleInput.val() || '弹窗标题',
      width: parseInt(this.widthInput.val() || 400, 10),
      height: parseInt(this.heightInput.val() || 300, 10),
      content: this.htmlTextarea.val() || ''
    };
  }
  
  // 保存弹窗数据
  save() {
    if (!this.currentCell || this.currentRi === null || this.currentCi === null) {
      console.error('没有当前编辑的单元格信息');
      return;
    }
    
    const popupData = this.getPopupData();
    
    // 调用回调函数保存数据
    this.change({
      ri: this.currentRi,
      ci: this.currentCi,
      cell: this.currentCell,
      popupData: popupData
    });
    
    this.hide();
  }
  
  // 设置位置
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
  
  // 设置当前编辑的单元格信息
  setCell(ri, ci, cell) {
    this.currentRi = ri;
    this.currentCi = ci;
    this.currentCell = cell;
    
    // 设置数据
    this.setPopupData(cell.popupData);
    
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
}