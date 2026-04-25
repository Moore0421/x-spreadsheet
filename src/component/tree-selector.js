import { h } from './element';
import { cssPrefix } from '../config';
import { bindClickoutside, unbindClickoutside } from './event';

export default class TreeSelector {
  constructor(options = {}) {
    this.options = options;
    this.items = options.items || [];
    this.searchText = '';
    
    // 创建主容器
    this.el = h('div', `${cssPrefix}-tree-selector`)
      .css('width', '300px')
      .css('height', '700px') // 将 max-height 改为固定 height，或者依然使用 max-height 并配合 flex
      .css('position', 'fixed')
      .css('background', '#fff')
      .css('box-shadow', '0 2px 12px rgba(0,0,0,0.15)')
      .css('border-radius', '4px')
      .css('z-index', '10000')
      .css('top', '50%')
      .css('left', '50%')
      .css('transform', 'translate(-50%, -50%)')
      .css('display', 'flex')
      .css('flex-direction', 'column')
      .hide();
    
    // 创建固定的顶部区域（包含标题和搜索框）
    this.fixedTop = h('div', `${cssPrefix}-tree-selector-fixed-top`)
      .css('flex', '0 0 auto'); // 确保顶部区域不伸缩，固定高度

    // 创建标题栏
    this.header = h('div', `${cssPrefix}-tree-selector-header`)
      .css('padding', '10px')
      .css('border-bottom', '1px solid #e0e0e0')
      .css('display', 'flex')
      .css('justify-content', 'space-between')
      .css('cursor', 'move')
      .css('align-items', 'center')
      .css('background', '#fff'); // 添加背景色，防止被内容穿透
    
    this.title = h('div', `${cssPrefix}-tree-selector-title`)
      .css('font-weight', 'bold')
      .html('选择树形项');
    
    this.closeBtn = h('div', `${cssPrefix}-tree-selector-close`)
      .css('cursor', 'pointer')
      .css('font-size', '16px')
      .html('×')
      .on('click', () => this.hide());
    
    this.header.children(this.title, this.closeBtn);
    
    // 创建搜索框区域
    this.searchWrapper = h('div', `${cssPrefix}-tree-selector-search-wrapper`)
      .css('padding', '10px')
      .css('border-bottom', '1px solid #e0e0e0')
      .css('background', '#fff'); // 添加背景色

    this.searchInput = h('input', `${cssPrefix}-tree-selector-search`)
      .attr('type', 'text')
      .attr('placeholder', '搜索...')
      .css('width', '100%')
      .css('box-sizing', 'border-box')
      .css('padding', '6px 10px')
      .css('border', '1px solid #dcdfe6')
      .css('border-radius', '4px')
      .css('outline', 'none')
      .on('input', (e) => {
        this.searchText = e.target.value.trim().toLowerCase();
        this.renderTree();
      });

    this.searchWrapper.child(this.searchInput);
    
    // 将标题和搜索框加入到固定顶部容器
    this.fixedTop.children(this.header, this.searchWrapper);

    // 创建内容区域（可滚动）
    this.content = h('div', `${cssPrefix}-tree-selector-content`)
      .css('padding', '10px')
      .css('overflow-y', 'auto')
      .css('flex', '1 1 auto'); // 让内容区域自适应剩余空间并可滚动
    
    // 创建树形结构容器
    this.tree = h('ul', `${cssPrefix}-tree-list`)
      .css('list-style', 'none')
      .css('padding-left', '0')
      .css('margin', '0');
    
    this.content.child(this.tree);
    this.el.children(this.fixedTop, this.content);
    
    // 绑定点击外部事件
    bindClickoutside(this.el, () => {
      if (this.el.el.style.display !== 'none') {
        this.hide();
      }
    });
    
    // 初始化树形结构
    this.renderTree();

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
  
  // 渲染树形结构
  renderTree() {
    const buildTreeNode = (items) => {
      let hasVisibleNode = false;
      const elements = [];

      items.forEach((item) => {
        // 如果有搜索词，判断当前节点或其子节点是否匹配
        let isMatch = true;
        let childrenElements = [];
        let hasVisibleChild = false;
        
        if (item.children && item.children.length > 0) {
          const result = buildTreeNode(item.children);
          childrenElements = result.elements;
          hasVisibleChild = result.hasVisibleNode;
        }

        if (this.searchText) {
          const matchLabel = item.label && item.label.toLowerCase().includes(this.searchText);
          isMatch = matchLabel || hasVisibleChild;
        }

        if (!isMatch) {
          return;
        }
        
        hasVisibleNode = true;

        const itemEl = h('li', `${cssPrefix}-tree-item`)
          .css('padding', '0')
          .css('cursor', 'pointer')
          .css('list-style', 'none');
        
        const itemRow = h('div', `${cssPrefix}-tree-item-row`)
          .css('display', 'flex')
          .css('align-items', 'center')
          .css('width', '100%')
          .css('padding', '8px 5px')
          .on('mouseenter', () => {
            itemRow.css('background', '#f0f0f0');
          })
          .on('mouseleave', () => {
            itemRow.css('background', 'transparent');
          });
        
        itemEl.child(itemRow);
        
        let toggleEl = null;
        let subList = null;
        
        if (item.children && item.children.length > 0) {
          // 如果正在搜索且匹配到了子节点，或者没有搜索，则根据原来的逻辑判断
          // 搜索状态下默认展开所有匹配的父节点
          const isExpanded = this.searchText ? true : false;

          toggleEl = h('span', `${cssPrefix}-tree-item-toggle`)
            .html(isExpanded ? '▼' : '▶')
            .css('margin-right', '5px')
            .css('cursor', 'pointer')
            .css('display', 'inline-flex')
            .css('align-items', 'center')
            .css('width', '15px')
            .css('font-size', '14px')
            .on('click', (e) => {
              e.stopPropagation();
              if (subList && subList.el.style.display === 'none') {
                subList.show();
                toggleEl.html('▼');
              } else if (subList) {
                subList.hide();
                toggleEl.html('▶');
              }
            });
          
          itemRow.child(toggleEl);
        } else {
          const spacer = h('span', '')
            .css('width', '15px')
            .css('margin-right', '5px');
          itemRow.child(spacer);
        }
        
        // 搜索关键词高亮
        let labelHtml = item.label;
        if (this.searchText && item.label && item.label.toLowerCase().includes(this.searchText)) {
          const regex = new RegExp(`(${this.searchText})`, 'gi');
          labelHtml = item.label.replace(regex, '<span style="color: #409eff; font-weight: bold;">$1</span>');
        }

        const labelEl = h('span', `${cssPrefix}-tree-item-label`)
          .html(labelHtml)
          .css('flex', '1')
          .css('font-size', '14px')
          .css('cursor', 'pointer')
          .on('click', (e) => {
            e.stopPropagation();
            this.change(item);
            this.hide();
          });
        
        itemRow.child(labelEl);
        
        itemRow.on('click', (e) => {
          if (!toggleEl || !toggleEl.el.contains(e.target)) {
            this.change(item);
            this.hide();
          }
        });
        
        if (item.children && item.children.length > 0) {
          subList = h('ul', `${cssPrefix}-tree-sublist`)
            .css('padding-left', '20px')
            .css('margin', '0');
          
          if (!this.searchText) {
             subList.hide();
          }
          
          subList.children(...childrenElements);
          itemEl.child(subList);
        }
        
        elements.push(itemEl);
      });

      return { elements, hasVisibleNode };
    };
    
    this.tree.html('');
    if (this.items && this.items.length > 0) {
      const result = buildTreeNode(this.items);
      if (result.elements.length > 0) {
        this.tree.children(...result.elements);
      } else {
        const emptyTip = h('div', `${cssPrefix}-tree-empty`)
          .css('padding', '20px')
          .css('text-align', 'center')
          .css('color', '#909399')
          .css('font-size', '14px')
          .html('无匹配数据');
        this.tree.child(emptyTip);
      }
    }
  }
  
  // 设置树形数据
  setItems(items) {
    this.items = items;
    this.renderTree();
  }
  
  // 设置回调函数
  setChange(callback) {
    this.change = callback;
    return this;
  }
  
  // 显示树形选择器
  show() {
    // 每次打开时清空搜索框内容
    this.searchText = '';
    this.searchInput.val('');
    this.renderTree();
    
    this.el.css('display', 'flex');
    return this;
  }
  
  // 隐藏树形选择器
  hide() {
    this.el.hide();
    return this;
  }
} 