import {
  CELL_RANGE_REGEX,
  CELL_REF_REGEX,
  CIRCULAR_DEPENDENCY_ERROR,
  DYNAMIC_VARIABLE_ERROR,
  DYNAMIC_VARIABLE_RESOLVING,
  GENERAL_ERROR,
  REF_ERROR,
  SHEET_TO_CELL_REF_REGEX,
  SPACE_REMOVAL_REGEX,
} from "../constants";
import { expr2xy, xy2expr } from "./alphabet";
import { numberCalc } from "./helper";
import { Parser } from "hot-formula-parser";
import dynamicFunctions from './dynamic_functions';

// Converting infix expression to a suffix expression
// src: AVERAGE(SUM(A1,A2), B1) + 50 + B20
// return: [A1, A2], SUM[, B1],AVERAGE,50,+,B20,+
const infixExprToSuffixExpr = (src) => {
  const operatorStack = [];
  const stack = [];
  let subStrs = []; // SUM, A1, B2, 50 ...
  let fnArgType = 0; // 1 => , 2 => :
  let fnArgOperator = "";
  let fnArgsLen = 1; // A1,A2,A3...
  let oldc = "";
  for (let i = 0; i < src.length; i += 1) {
    const c = src.charAt(i);
    if (c !== " ") {
      if (c >= "a" && c <= "z") {
        subStrs.push(c.toUpperCase());
      } else if (
        (c >= "0" && c <= "9") ||
        (c >= "A" && c <= "Z") ||
        c === "."
      ) {
        subStrs.push(c);
      } else if (c === '"') {
        i += 1;
        while (src.charAt(i) !== '"') {
          subStrs.push(src.charAt(i));
          i += 1;
        }
        stack.push(`"${subStrs.join("")}`);
        subStrs = [];
      } else if (c === "-" && /[+\-*/,(]/.test(oldc)) {
        subStrs.push(c);
      } else {
        // console.log('subStrs:', subStrs.join(''), stack);
        if (c !== "(" && subStrs.length > 0) {
          stack.push(subStrs.join(""));
        }
        if (c === ")") {
          let c1 = operatorStack.pop();
          if (fnArgType === 2) {
            // fn argument range => A1:B5
            try {
              const [ex, ey] = expr2xy(stack.pop());
              const [sx, sy] = expr2xy(stack.pop());
              // console.log('::', sx, sy, ex, ey);
              let rangelen = 0;
              for (let x = sx; x <= ex; x += 1) {
                for (let y = sy; y <= ey; y += 1) {
                  stack.push(xy2expr(x, y));
                  rangelen += 1;
                }
              }
              stack.push([c1, rangelen]);
            } catch (e) {
              // console.log(e);
            }
          } else if (fnArgType === 1 || fnArgType === 3) {
            if (fnArgType === 3) stack.push(fnArgOperator);
            // fn argument => A1,A2,B5
            stack.push([c1, fnArgsLen]);
            fnArgsLen = 1;
          } else {
            // console.log('c1:', c1, fnArgType, stack, operatorStack);
            while (c1 !== "(") {
              stack.push(c1);
              if (operatorStack.length <= 0) break;
              c1 = operatorStack.pop();
            }
          }
          fnArgType = 0;
        } else if (c === "=" || c === ">" || c === "<") {
          const nc = src.charAt(i + 1);
          fnArgOperator = c;
          if (nc === "=" || nc === "-") {
            fnArgOperator += nc;
            i += 1;
          }
          fnArgType = 3;
        } else if (c === ":") {
          fnArgType = 2;
        } else if (c === ",") {
          if (fnArgType === 3) {
            stack.push(fnArgOperator);
          }
          fnArgType = 1;
          fnArgsLen += 1;
        } else if (c === "(" && subStrs.length > 0) {
          // function
          operatorStack.push(subStrs.join(""));
        } else {
          // priority: */ > +-
          // console.log('xxxx:', operatorStack, c, stack);
          if (operatorStack.length > 0 && (c === "+" || c === "-")) {
            let top = operatorStack[operatorStack.length - 1];
            if (top !== "(") stack.push(operatorStack.pop());
            if (top === "*" || top === "/") {
              while (operatorStack.length > 0) {
                top = operatorStack[operatorStack.length - 1];
                if (top !== "(") stack.push(operatorStack.pop());
                else break;
              }
            }
          } else if (operatorStack.length > 0) {
            const top = operatorStack[operatorStack.length - 1];
            if (top === "*" || top === "/") stack.push(operatorStack.pop());
          }
          operatorStack.push(c);
        }
        subStrs = [];
      }
      oldc = c;
    }
  }
  if (subStrs.length > 0) {
    stack.push(subStrs.join(""));
  }
  while (operatorStack.length > 0) {
    stack.push(operatorStack.pop());
  }
  return stack;
};

//This function returns value by converting into number and resolve cell value
const evalSubExpr = (subExpr, cellRender) => {
  const [fl] = subExpr;
  let expr = subExpr;
  if (fl === '"') {
    return subExpr.substring(1);
  }
  let ret = 1;
  if (fl === "-") {
    expr = subExpr.substring(1);
    ret = -1;
  }
  if (expr[0] >= "0" && expr[0] <= "9") {
    return ret * Number(expr);
  }
  const [x, y] = expr2xy(expr);
  return ret * cellRender(x, y);
};

// evaluate the suffix expression
// srcStack: <= infixExprToSufixExpr
// formulaMap: {'SUM': {}, ...}
// cellRender: (x, y) => {}
const evalSuffixExpr = (srcStack, formulaMap, cellRender, cellList) => {
  const stack = [];
  // console.log(':::::formulaMap:', formulaMap);
  for (let i = 0; i < srcStack.length; i += 1) {
    // console.log(':::>>>', srcStack[i]);
    const expr = srcStack[i];
    const fc = expr[0];
    if (expr === "+") {
      const top = stack.pop();
      stack.push(numberCalc("+", stack.pop(), top));
    } else if (expr === "-") {
      if (stack.length === 1) {
        const top = stack.pop();
        stack.push(numberCalc("*", top, -1));
      } else {
        const top = stack.pop();
        stack.push(numberCalc("-", stack.pop(), top));
      }
    } else if (expr === "*") {
      stack.push(numberCalc("*", stack.pop(), stack.pop()));
    } else if (expr === "/") {
      const top = stack.pop();
      stack.push(numberCalc("/", stack.pop(), top));
    } else if (expr === "^") {
      const top = stack.pop();
      stack.push(numberCalc("^", stack.pop(), top));
    } else if (fc === "=" || fc === ">" || fc === "<") {
      let top = stack.pop();
      if (!Number.isNaN(top)) top = Number(top);
      let left = stack.pop();
      if (!Number.isNaN(left)) left = Number(left);
      let ret = false;
      if (fc === "=") {
        ret = left === top;
      } else if (expr === ">") {
        ret = left > top;
      } else if (expr === ">=") {
        ret = left >= top;
      } else if (expr === "<") {
        ret = left < top;
      } else if (expr === "<=") {
        ret = left <= top;
      }
      stack.push(ret);
    } else if (fc === '$') {
      const funcName = expr.split('.')[0].substring(1);
      const funcPath = expr.substring(funcName.length + 1);
      const func = dynamicFunctions.getFunction(funcName);
      if (!func) {
        return `#UNDEFINED_FUNCTION!`;
      }
      try {
        const result = func();
        // 处理对象属性访问（如 people[0].name）
        const value = funcPath.split('.').reduce((acc, cur) => {
          const [prop, index] = cur.split('[');
          if (index !== undefined) {
            return acc[prop][index.replace(']', '')];
          }
          return acc[prop] ?? `#PROPERTY_NOT_FOUND!`;
        }, result);
        stack.push(value);
      } catch (e) {
        return `#EVALUATION_ERROR!`;
      }
    } else if (Array.isArray(expr)) {
      const [formula, len] = expr;
      const params = [];
      for (let j = 0; j < len; j += 1) {
        params.push(stack.pop());
      }
      if (formulaMap[formula]) {
        stack.push(formulaMap[formula].render(params.reverse()));
      }
    } else {
      if (cellList.includes(expr)) {
        return 0;
      }
      if ((fc >= "a" && fc <= "z") || (fc >= "A" && fc <= "Z")) {
        cellList.push(expr);
      }
      stack.push(evalSubExpr(expr, cellRender));
      cellList.pop();
    }
    // console.log('stack:', stack);
  }
  return stack[0];
};

const rangeToCellConversion = (range) => {
  let cells = "";
  if (range.length) {
    const cellInfo = range?.toUpperCase()?.split(":");
    if (cellInfo.length === 2) {
      const [ex, ey] = expr2xy(cellInfo[1]);
      const [sx, sy] = expr2xy(cellInfo[0]);
      const cellArray = [];
      for (let x = sx; x <= ex; x += 1) {
        for (let y = sy; y <= ey; y += 1) {
          cellArray.push(xy2expr(x, y));
        }
      }
      return cellArray?.join(",");
    }
  }
  return cells;
};

const parserFormulaString = (
  string,
  getCellText,
  cellRender,
  getDynamicVariable,
  trigger,
  formulaCallStack,
  sheetName
) => {
  if (string?.length) {
    try {
      let isFormulaResolved = false;
      let newFormulaString = string;
      let dynamicVariableError = false;
      let isCircularDependency = false;
      let isVariableResolving = false;
      if (trigger) {
        let dynamicVariableRegEx = new RegExp(`\\${trigger}\\S*`, "g");
        newFormulaString = newFormulaString.replace(
          dynamicVariableRegEx,
          (match) => {
            const { text, resolved, resolving } = getDynamicVariable(match);
            if (resolving) isVariableResolving = true;
            else if (resolved) return text;
            else dynamicVariableError = true;
          }
        );
      }
      if (isVariableResolving) return DYNAMIC_VARIABLE_RESOLVING;
      else if (dynamicVariableError) return DYNAMIC_VARIABLE_ERROR;
      // Removing spaces other than the spaces that are in apostrophes
      newFormulaString = newFormulaString.replace(SPACE_REMOVAL_REGEX, "");
      newFormulaString = newFormulaString.replace(
        SHEET_TO_CELL_REF_REGEX,
        (match) => {
          const [linkSheetName, cellRef] = match.replaceAll("'", "").split("!");
          const [x, y] = expr2xy(cellRef);
          const text = getCellText(x, y, linkSheetName);
          if (text?.startsWith?.("=")) {
            if (formulaCallStack?.[linkSheetName]?.includes(cellRef))
              isCircularDependency = true;
            else {
              formulaCallStack[linkSheetName] =
                formulaCallStack[linkSheetName] || [];
              formulaCallStack[linkSheetName].push(cellRef);
            }
            return isCircularDependency
              ? 0
              : cellRender(
                  text,
                  getCellText,
                  getDynamicVariable,
                  trigger,
                  formulaCallStack,
                  linkSheetName
                );
          }
          if (text === REF_ERROR) isFormulaResolved = true;
          return isNaN(Number(text)) ? `"${text}"` : text;
        }
      );

      if (isFormulaResolved) return REF_ERROR;
      newFormulaString = newFormulaString.replace(CELL_RANGE_REGEX, (match) => {
        const cells = rangeToCellConversion(match);
        if (cells?.length) {
          return cells;
        }
      });
      newFormulaString = newFormulaString.replace(CELL_REF_REGEX, (cellRef) => {
        const [x, y] = expr2xy(cellRef);
        const text = getCellText(x, y);
        if (text) {
          if (text?.startsWith?.("=")) {
            if (formulaCallStack?.[sheetName]?.includes(cellRef))
              isCircularDependency = true;
            else {
              formulaCallStack[sheetName] = formulaCallStack[sheetName] || [];
              formulaCallStack[sheetName].push(cellRef);
            }
            return isCircularDependency
              ? 0
              : cellRender(
                  text,
                  getCellText,
                  getDynamicVariable,
                  trigger,
                  formulaCallStack,
                  sheetName
                );
          } else {
            return isNaN(Number(text)) ? `"${text}"` : text;
          }
        } else {
          return 0;
        }
      });
      return isCircularDependency
        ? CIRCULAR_DEPENDENCY_ERROR
        : newFormulaString;
    } catch (e) {
      return string;
    }
  }
  return string;
};

const cellRender = (
  src,
  getCellText,
  getDynamicVariable,
  trigger,
  formulaCallStack = {},
  sheetName
) => {
  if (src[0] === "=") {
    const formula = src.substring(1);
    try {
      var parser = new Parser();
      const parsedFormula = parserFormulaString(
        formula,
        getCellText,
        cellRender,
        getDynamicVariable,
        trigger,
        formulaCallStack,
        sheetName
      );

      if (parsedFormula.includes(REF_ERROR)) return REF_ERROR;
      else if (parsedFormula.includes(CIRCULAR_DEPENDENCY_ERROR))
        return CIRCULAR_DEPENDENCY_ERROR;
      else if (parsedFormula.includes(DYNAMIC_VARIABLE_RESOLVING))
        return DYNAMIC_VARIABLE_RESOLVING;
      else if (parsedFormula.includes(DYNAMIC_VARIABLE_ERROR))
        return DYNAMIC_VARIABLE_ERROR;
      const data = parser.parse(parsedFormula);
      return data?.error?.replace("#", "") ?? data?.result;
    } catch (e) {
      return GENERAL_ERROR;
    }
  }
  
  // 处理JavaScript代码（$$前缀）
  if (src.startsWith("$$")) {
    // 直接在cellRender中执行JavaScript代码并返回结果
    try {
      const jsCode = src.substring(2);
      console.log("cellRender执行JS代码:", jsCode);
      
      // 直接执行JavaScript代码并返回结果
      return executeJavaScriptInSandbox(jsCode);
    } catch (e) {
      console.error("JS执行错误:", e);
      return `#JS_ERROR! ${e.message}`;
    }
  }
  
  // 处理以$开头的动态函数表达式
  if (src[0] === "$" && !src.startsWith("$$")) {
    // 检查是否包含运算符
    const operators = ['+', '-', '*', '/', '===', '==', '!==', '!=', '>=', '<=', '>', '<', '&&', '||'];
    const hasOperator = operators.some(op => src.includes(op));
    
    if (hasOperator) {
      // 是JS运算表达式，交给evaluateJsExpression处理
      try {
        return evaluateJsExpression(src.substring(1));
      } catch (e) {
        return `#JS_ERROR: ${e.message}`;
      }
    } else {
      // 是动态函数表达式
      try {
        return evaluateDynamicExpression(src.substring(1));
      } catch (e) {
        return `#ERROR: ${e.message}`;
      }
    }
  }
  
  if (src[0] === trigger) {
    const { text, resolved, resolving } = getDynamicVariable(src);
    return resolving
      ? DYNAMIC_VARIABLE_RESOLVING
      : resolved
        ? text ?? src
        : DYNAMIC_VARIABLE_ERROR;
  }
  
  return src;
};

// 更新动态表达式评估函数，添加保护检查
const evaluateDynamicExpression = (expression, dataProxy) => {
  // 添加保护检查，防止$$代码被错误传入
  if (expression.startsWith('$')) {
    console.warn('错误：$$前缀代码被错误地传递给了evaluateDynamicExpression');
    return "#ROUTING_ERROR! JavaScript代码应使用executeJavaScriptInSandbox处理";
  }
  
  if (!expression) return "#SYNTAX_ERROR! expression is empty";
  
  try {
    // 检查是否是运算表达式（包含运算符）
    const operators = ['+', '-', '*', '/', '===', '==', '!==', '!=', '>=', '<=', '>', '<', '&&', '||'];
    const hasOperator = operators.some(op => expression.includes(op));
    
    // 如果是运算表达式，则进行特殊处理
    if (hasOperator) {
      return evaluateJsExpression(expression, dataProxy);
    }
    
    // 原始函数解析逻辑保持不变
    // 匹配模式：functionName(arg1,arg2,...) 或 functionName
    const functionMatch = expression.match(/^([a-zA-Z0-9_]+)(?:\((.*)\))?/);
    
    if (!functionMatch) {
      return "#SYNTAX_ERROR! Invalid function format";
    }
    
    const functionName = functionMatch[1];
    const argsString = functionMatch[2] || '';
    
    // 解析参数（支持简单参数，未处理嵌套括号和引号内的逗号）
    const args = argsString ? argsString.split(',').map(arg => {
      const trimmed = arg.trim();
      // 尝试将数字字符串转换为数字
      if (/^-?\d+$/.test(trimmed)) {
        return parseInt(trimmed, 10);
      } else if (/^-?\d+\.\d+$/.test(trimmed)) {
        return parseFloat(trimmed);
      } else if (trimmed === 'true') {
        return true;
      } else if (trimmed === 'false') {
        return false;
      } else if (trimmed === 'null') {
        return null;
      } else if (trimmed === 'undefined') {
        return undefined;
      } else if (/^["'].*["']$/.test(trimmed)) {
        // 处理引号中的字符串
        return trimmed.slice(1, -1);
      }
      return trimmed;
    }) : [];
    
    // 执行函数获取结果
    let result;
    try {
      // 使用dynamicFunctions模块执行函数并传递参数
      if (!dynamicFunctions.getFunction(functionName)) {
        // 尝试进行不区分大小写的查找（如果需要）
        const allFunctions = Object.keys(dynamicFunctions.listFunctions());
        const functionNameLower = functionName.toLowerCase();
        const matchedFunction = allFunctions.find(f => f.toLowerCase() === functionNameLower);
        
        if (matchedFunction) {
          // console.log(`找到不区分大小写的函数匹配: ${functionName} -> ${matchedFunction}`);
          result = dynamicFunctions.executeFunction(matchedFunction, args);
        } else {
          return `#UNDEFINED_FUNCTION!`;
        }
      } else {
        result = dynamicFunctions.executeFunction(functionName, args);
      }
    } catch (e) {
      return `#FUNCTION_ERROR! ${e.message}`;
    }
    
    // 检查结果是否为null或undefined
    if (result === null) {
      return "#NULL!";
    }
    
    if (result === undefined) {
      return "#UNDEFINED!";
    }
    
    // 处理属性访问路径（如果有）
    const pathPart = expression.substring(functionMatch[0].length);
    
    if (pathPart) {
      // 获取第一个点之后的访问路径
      const accessPath = pathPart.startsWith('.') ? pathPart.substring(1) : pathPart;
      const parts = accessPath.split('.');
      
      // 解析属性访问路径
      for (let i = 0; i < parts.length; i++) {
        let part = parts[i];
        
        // 处理数组索引访问，例如: people[0]
        if (part.includes('[')) {
          const arrayAccessParts = part.split('[');
          part = arrayAccessParts[0];
          
          // 可能有多个连续的数组访问，例如: matrix[0][1]
          if (part) {
            if (typeof result !== 'object' || result === null) {
              return `#TYPE_ERROR!`;
            }
            result = result[part];
            if (result === undefined) {
              return `#PROPERTY_NOT_FOUND!`;
            }
            if (result === null) {
              return `#NULL!`;
            }
          }
          
          // 处理所有的数组索引
          for (let j = 1; j < arrayAccessParts.length; j++) {
            const indexStr = arrayAccessParts[j].replace(']', '');
            
            // 检查索引是否为有效数字
            if (!/^\d+$/.test(indexStr)) {
              return `#SYNTAX_ERROR! indexStr is not a number`;
            }
            
            const index = parseInt(indexStr, 10);
            
            // 检查是否为数组
            if (!Array.isArray(result)) {
              return `#TYPE_ERROR!`;
            }
            
            // 检查索引是否越界
            if (index < 0 || index >= result.length) {
              return `#INDEX_OUT_OF_RANGE!`;
            }
            
            result = result[index];
            
            // 检查结果是否为null或undefined
            if (result === null) {
              return `#NULL!`;
            }
            if (result === undefined) {
              return `#UNDEFINED!`;
            }
          }
        } else {
          // 简单属性访问
          if (typeof result !== 'object' || result === null) {
            return `#TYPE_ERROR!`;
          }
          
          result = result[part];
          
          // 检查路径中的undefined值
          if (result === undefined) {
            return `#PROPERTY_NOT_FOUND!`;
          }
          
          // 检查null值
          if (result === null) {
            return `#NULL!`;
          }
        }
      }
    }
    
    // 处理最终结果
    if (typeof result === 'object' && result !== null) {
      // 对象和数组转为JSON字符串
      return JSON.stringify(result);
    }
    
    return result?.toString() || '';
  } catch (e) {
    return `#EVALUATION_ERROR! ${e.message}`;
  }
};

// 新增JavaScript表达式求值函数
const evaluateJsExpression = (expression, dataProxy) => {
  try {
    // console.log('计算JS表达式:', expression);
    
    // 1. 查找表达式中的所有动态函数引用
    const functionRefs = new Set();
    const dynamicFuncRegex = /\b([a-zA-Z0-9_]+)(?:\([^)]*\))?/g;
    let match;
    
    while ((match = dynamicFuncRegex.exec(expression)) !== null) {
      const funcName = match[1];
      // 检查是否是已注册的动态函数
      if (dynamicFunctions.getFunction(funcName)) {
        functionRefs.add(funcName);
      }
    }
    
    // 2. 预处理表达式，先解析所有的动态函数调用
    let processedExpr = expression;
    
    // 匹配所有函数调用和属性访问模式
    const fullExpressionRegex = /\b([a-zA-Z0-9_]+)(\([^)]*\))?(\.[a-zA-Z0-9_]+|\[[^\]]+\])*/g;
    
    // 创建一个临时变量存储函数结果
    const tempVars = {};
    let varCounter = 0;
    
    // 替换所有复杂表达式为临时变量
    processedExpr = processedExpr.replace(fullExpressionRegex, (fullMatch, funcName) => {
      // 如果这不是一个动态函数调用，保持原样
      if (!functionRefs.has(funcName)) {
        return fullMatch;
      }
      
      const varName = `__temp${varCounter++}`;
      
      try {
        // 评估动态函数表达式
        const result = evaluateDynamicExpression(fullMatch, dataProxy);
        tempVars[varName] = result;
        return varName;
      } catch (e) {
        console.error(`动态函数 ${funcName} 执行错误:`, e);
        // 发生错误时将其记录并保持表达式原样
        return fullMatch;
      }
    });
    
    // console.log('处理后的表达式:', processedExpr);
    // console.log('临时变量:', tempVars);
    
    // 3. 构建上下文对象，包含动态函数和临时变量
    const context = {
      ...tempVars,
    };
    
    // 为每个引用的动态函数添加到上下文
    functionRefs.forEach(funcName => {
      const func = dynamicFunctions.getFunction(funcName);
      if (func) {
        context[funcName] = (...args) => {
          // console.log(`JS表达式中调用函数 ${funcName}:`, args);
          return dynamicFunctions.executeFunction(funcName, args);
        };
      }
    });
    
    // 4. 构建计算函数
    const argNames = Object.keys(context);
    const argValues = Object.values(context);
    
    // console.log('表达式上下文变量:', argNames);
    
    // 创建并执行函数
    const calcFunc = new Function(...argNames, `
      try {
        return ${processedExpr};
      } catch (e) {
        return "#CALCULATION_ERROR! " + e.message;
      }
    `);
    
    // 执行计算
    let result = calcFunc(...argValues);
    
    // 5. 格式化结果
    if (result === true) return "TRUE";
    if (result === false) return "FALSE";
    if (result === null) return "#NULL!";
    if (result === undefined) return "#UNDEFINED!";
    
    return result;
  } catch (e) {
    console.error('JS表达式计算错误:', e);
    return `#JS_EVALUATION_ERROR! ${e.message}`;
  }
};

/**
 * 在安全沙箱中执行JavaScript代码并获取返回值（同步版本）
 * @param {string} code 要执行的JavaScript代码
 * @returns {*} 执行结果或错误信息
 */
const executeJavaScriptInSandbox = (code) => {
  try {
    console.log('执行JavaScript代码:', code);
    
    // 处理代码字符串，确保能够正确捕获return值
    let processedCode = code.trim();
    
    // 超时保护 - 记录开始时间
    const startTime = Date.now();
    const MAX_EXECUTION_TIME = 5000; // 5秒超时限制
    let isTimedOut = false;
    
    // 检查代码类型，并准备适当的执行方式
    const hasExplicitReturn = /\breturn\b/.test(processedCode);
    const isSimpleExpression = !processedCode.includes(';') && 
                             !processedCode.includes('{') && 
                             !processedCode.includes('\n');
    
    // 检测潜在的无限循环
    const hasLoops = /\b(for|while|do)\b/.test(processedCode);
    const hasRecursion = /function\s+\w+\s*\([^)]*\)\s*{[^}]*\1\s*\(/.test(processedCode);
    
    if (hasLoops || hasRecursion) {
      console.warn('检测到可能的循环或递归，添加超时保护');
    }
    
    let executableCode;
    if (isSimpleExpression && !hasExplicitReturn) {
      // 简单表达式，自动添加return
      executableCode = `return (${processedCode});`;
    } else if (hasExplicitReturn) {
      // 已有return语句，确保在函数体内执行
      executableCode = processedCode;
    } else {
      // 复杂代码，可能隐式返回undefined
      executableCode = processedCode;
    }
    
    console.log('处理后的代码:', executableCode);
    
    // 创建沙箱环境变量
    const sandbox = {
      // 基本类型和构造函数
      Object, Array, String, Number, Boolean, Date, RegExp, Math, JSON,
      
      // 添加超时检查函数
      __checkTimeout: function() {
        if (Date.now() - startTime > MAX_EXECUTION_TIME) {
          isTimedOut = true;
          throw new Error('代码执行超时(5秒)');
        }
      },
      
      // 安全的控制台方法
      console: {
        log: (...args) => {
          console.log('沙箱控制台:', ...args);
          sandbox.__checkTimeout(); // 每次调用检查超时
          return undefined;
        },
        warn: (...args) => {
          console.warn('沙箱控制台:', ...args);
          sandbox.__checkTimeout();
          return undefined;
        },
        error: (...args) => {
          console.error('沙箱控制台:', ...args);
          sandbox.__checkTimeout();
          return undefined;
        }
      },
      
      // 工具函数 - 添加超时检查
      setTimeout: (fn, delay) => {
        sandbox.__checkTimeout();
        const safeDelay = Math.min(delay || 0, 1000); // 限制最大延迟
        return setTimeout(() => {
          sandbox.__checkTimeout();
          fn();
        }, safeDelay);
      },
      
      // 重写数组方法，添加超时检查
      Array: {
        ...Array,
        prototype: {
          ...Array.prototype,
          forEach: function(callback, thisArg) {
            sandbox.__checkTimeout();
            return Array.prototype.forEach.call(this, (...args) => {
              sandbox.__checkTimeout();
              return callback.apply(thisArg, args);
            });
          },
          map: function(callback, thisArg) {
            sandbox.__checkTimeout();
            return Array.prototype.map.call(this, (...args) => {
              sandbox.__checkTimeout();
              return callback.apply(thisArg, args);
            });
          },
          filter: function(callback, thisArg) {
            sandbox.__checkTimeout();
            return Array.prototype.filter.call(this, (...args) => {
              sandbox.__checkTimeout();
              return callback.apply(thisArg, args);
            });
          },
          reduce: function(callback, initialValue) {
            sandbox.__checkTimeout();
            return Array.prototype.reduce.call(this, (...args) => {
              sandbox.__checkTimeout();
              return callback.apply(null, args);
            }, initialValue);
          }
        }
      },
      
      // 数据处理工具
      parseInt, parseFloat, isNaN, isFinite,
      encodeURI, decodeURI, encodeURIComponent, decodeURIComponent,
      
      // 自定义工具函数
      utils: {
        formatNumber: (num, decimals = 2) => {
          sandbox.__checkTimeout();
          return Number(num).toFixed(decimals);
        },
        formatDate: (date, format = 'YYYY-MM-DD') => {
          sandbox.__checkTimeout();
          const d = new Date(date);
          if (isNaN(d.getTime())) return 'Invalid Date';
          
          const year = d.getFullYear();
          const month = String(d.getMonth() + 1).padStart(2, '0');
          const day = String(d.getDate()).padStart(2, '0');
          
          return format
            .replace('YYYY', year)
            .replace('MM', month)
            .replace('DD', day);
        },
        sum: (array) => {
          sandbox.__checkTimeout();
          if (!Array.isArray(array)) return 0;
          let result = 0;
          for (let i = 0; i < array.length; i++) {
            sandbox.__checkTimeout();
            result += (Number(array[i]) || 0);
          }
          return result;
        },
        avg: (array) => {
          sandbox.__checkTimeout();
          if (!Array.isArray(array) || array.length === 0) return 0;
          return sandbox.utils.sum(array) / array.length;
        }
      },
      
      // 禁止访问的对象设为null
      window: null,
      document: null,
      navigator: null,
      location: null,
      fetch: null,
      XMLHttpRequest: null
    };
    
    // 如果有循环或递归，修改代码注入超时检查
    if (hasLoops || hasRecursion) {
      // 在循环和函数调用前插入超时检查
      executableCode = executableCode
        .replace(/(\b(for|while|do)\b[^{]*{)/g, '$1 __checkTimeout();')
        .replace(/(\bfunction\b[^{]*{)/g, '$1 __checkTimeout();');
    }
    
    // 构建函数参数列表
    const sandboxKeys = Object.keys(sandbox);
    const sandboxValues = Object.values(sandbox);
    
    // 创建函数封装并执行代码
    let sandboxFunction;
    
    try {
      if (isSimpleExpression && !hasExplicitReturn) {
        // 简单表达式，直接求值并返回
        sandboxFunction = new Function(...sandboxKeys, `
          "use strict";
          try {
            __checkTimeout(); // 初始检查
            const result = (${processedCode});
            __checkTimeout(); // 结果检查
            return result;
          } catch(e) {
            if (e.message.includes('超时')) {
              return "#TIMEOUT_ERROR! 代码执行超过5秒限制";
            }
            throw e;
          }
        `);
      } else {
        // 复杂代码块，包装在函数中
        sandboxFunction = new Function(...sandboxKeys, `
          "use strict";
          try {
            __checkTimeout(); // 初始检查
            let __result;
            
            // 包装在函数中来捕获return值
            __result = (function() {
              ${executableCode}
            })();
            
            __checkTimeout(); // 结果检查
            return __result;
          } catch(e) {
            if (e.message.includes('超时')) {
              return "#TIMEOUT_ERROR! 代码执行超过5秒限制";
            }
            throw e;
          }
        `);
      }
      
      // 执行函数并获取结果
      const result = sandboxFunction(...sandboxValues);
      
      // 再次检查是否超时
      if (isTimedOut) {
        return "#TIMEOUT_ERROR! 代码执行超过5秒限制";
      }
      
      // 格式化结果
      if (result === undefined || result === null) {
        return ''; // 没有返回值则显示空
      } else if (typeof result === 'object') {
        try {
          return JSON.stringify(result);
        } catch (e) {
          return String(result);
        }
      } else {
        return result;
      }
    } catch (e) {
      console.error('JavaScript执行错误:', e);
      return `#JS_ERROR! ${e.message}`;
    }
  } catch (e) {
    console.error('沙箱创建错误:', e);
    return `#JS_SANDBOX_ERROR! ${e.message}`;
  }
};

export default {
  render: cellRender,
};
export { infixExprToSuffixExpr, evaluateDynamicExpression };
