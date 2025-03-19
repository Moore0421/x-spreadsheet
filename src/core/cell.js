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

    //Commented This functionality of formula calculation on X-Spread sheet and doing it by our own way

    // const stack = infixExprToSuffixExpr(src.substring(1));
    // console.log(stack, "stack");
    // if (stack.length <= 0) return src;
    // return evalSuffixExpr(
    //   stack,
    //   formulaMap,
    //   (x, y) =>
    //     cellRender(getCellText(x, y), formulaMap, getCellText, cellList),
    //   cellList
    // );
  }
  
  // 处理以$开头的动态函数表达式
  if (src[0] === "$") {
    const expression = src.substring(1);
    try {
      return evaluateDynamicExpression(expression);
    } catch (e) {
      return `#ERROR: ${e.message}`;
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

// 添加动态表达式评估函数
const evaluateDynamicExpression = (expression) => {
  if (!expression) return "#SYNTAX_ERROR!";
  
  // 匹配函数名和访问路径
  const parts = expression.split('.');
  const functionName = parts[0];
  
  try {
    // 执行函数获取结果
    let result;
    try {
      // 使用dynamicFunctions模块执行函数
      if (!dynamicFunctions.getFunction(functionName)) {
        return `#UNDEFINED_FUNCTION!`;
      }
      result = dynamicFunctions.executeFunction(functionName);
    } catch (e) {
      return `#FUNCTION_ERROR!`;
    }
    
    // 检查结果是否为null或undefined
    if (result === null) {
      return "#NULL!";
    }
    
    if (result === undefined) {
      return "#UNDEFINED!";
    }
    
    // 如果有后续访问路径，解析属性或数组
    if (parts.length > 1) {
      for (let i = 1; i < parts.length; i++) {
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
              return `#SYNTAX_ERROR!`;
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
    return `#EVALUATION_ERROR!`;
  }
};

export default {
  render: cellRender,
};
export { infixExprToSuffixExpr, evaluateDynamicExpression };
