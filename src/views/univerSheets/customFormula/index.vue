<template>
  <div class="univer-sheets-container">
    <div class="toolbar">
      <button @click="saveData" class="save-button">保存数据</button>
    </div>
    <div ref="container" class="univer-sheet"></div>
  </div>
</template>

<script lang="ts" setup>
import type { FUniver, Univer, Workbook, IWorkbookData } from "@univerjs/presets";
import type { BaseValueObject, IFunctionInfo, IFunctionNames } from "@univerjs/preset-sheets-core";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCoreZhCN from "@univerjs/preset-sheets-core/lib/locales/zh-CN.js";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";

// 自定义公式相关导入
import { 
  BaseFunction, 
  NumberValueObject, 
  FunctionType
} from '@univerjs/engine-formula';
import { onBeforeUnmount, onMounted, ref } from "vue";
import { ElMessage } from "element-plus";
import { saveSheetData } from "@/api/modules/univer";

import "@univerjs/preset-sheets-core/lib/index.css";

const container = ref<HTMLElement | null>(null);

// 定义自定义函数名称 - 使用常量而非对象
enum FUNCTION_NAMES_USER{
  CUSTOM_MULTIPLY_ADD = 'CUSTOM_MULTIPLY_ADD',
  CUSTOM_POWER_SUM = 'CUSTOM_POWER_SUM'
};

// 定义自定义函数实现
class CustomMultiplyAdd extends BaseFunction {
  override calculate(first: BaseValueObject, second: BaseValueObject, third: BaseValueObject) {
    // 实现 first * second + third
    const result = first.multiply(second).plus(third);
    return result;
  }
}

class CustomPowerSum extends BaseFunction {
  override calculate(...variants: BaseValueObject[]) {
    // 实现所有参数的平方和
    let accumulatorAll: BaseValueObject = new NumberValueObject(0);
    for (let i = 0; i < variants.length; i++) {
      let variant = variants[i];
      
      if (variant.isError()) {
        return variant;
      }
      
      // 计算平方
      const square = variant.multiply(variant);
      accumulatorAll = accumulatorAll.plus(square);
    }
    
    return accumulatorAll;
  }
}

const functionZhCN = {
  formula: {
    functionList: {
      CUSTOM_MULTIPLY_ADD: {
        description: '将单个值、单元格引用或是区域相加，或者将三者的组合相加。',
        abstract: '求积和和',
        links: [
          {
            title: '教学',
            url: 'https://support.microsoft.com/zh-cn/office/sum-%E5%87%BD%E6%95%B0-043e'
          }
        ],
        functionParameter: [
          {
            name: 'number1',
            detail: '要相乘的第一个数字或范围。',
            example: 'A1'
          },
          {
            name: 'number2',
            detail: '要相乘的第二个数字或范围。',
            example: 'B1'
          },
          {
            name: 'number3',
            detail: '要相加的第三个数字或范围。',
            example: 'C1'
          }
        ]
      },
      CUSTOM_POWER_SUM: {
        description: '求平方和',
        abstract: '平方和',
        links: [
          {
            title: '教学',
            url: 'https://univer.ai'
          }
        ],
        functionParameter: [
          {
            name: 'value',
            detail: '查询信息',
            example: 'A1'
          }
        ]
      }
    }
  }
}

// 定义函数信息
const FUNCTION_LIST_USER: IFunctionInfo[] = [
  {
    functionName: FUNCTION_NAMES_USER.CUSTOM_MULTIPLY_ADD,
    aliasFunctionName: "自定义乘法加法函数",
    functionType: FunctionType.Univer,
    description: "计算 a * b + c",
    abstract: "乘法加法运算",
    functionParameter: [
      {
        name: "第一个数值",
        detail: "参与运算的第一个数值",
        example: "A1",
        require: 1,
        repeat: 0
      },
      {
        name: "第二个数值",
        detail: "参与运算的第二个数值",
        example: "B1",
        require: 1,
        repeat: 0
      },
      {
        name: "第三个数值",
        detail: "参与运算的第三个数值",
        example: "C1",
        require: 1,
        repeat: 0
      }
    ]
  },
  {
    functionName: FUNCTION_NAMES_USER.CUSTOM_POWER_SUM,
    aliasFunctionName: "自定义平方和函数",
    functionType: FunctionType.Univer,
    description: "计算所有参数的平方和",
    abstract: "平方和运算",
    functionParameter: [
      {
        name: "数值1",
        detail: "参与运算的第一个数值",
        example: "A1",
        require: 1,
        repeat: 1
      }
    ]
  }
];

// 修改函数注册数组 - 使用正确的类型
const functionUser: Array<[new (name: IFunctionNames) => BaseFunction, IFunctionNames]> = [
  [CustomMultiplyAdd, FUNCTION_NAMES_USER.CUSTOM_MULTIPLY_ADD],
  [CustomPowerSum, FUNCTION_NAMES_USER.CUSTOM_POWER_SUM]
];

let univerInstance: Univer | null = null;
let univerAPIInstance: FUniver | null = null;
let workbook: Workbook | null = null;

onMounted(async() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(
        UniverPresetSheetsCoreZhCN,
        functionZhCN
      )
    },
    presets: [
      UniverSheetsCorePreset({
        container: container.value as HTMLElement,
        formula: {
          function: functionUser,
          description: FUNCTION_LIST_USER
        }
      })
    ]
  });

  workbook = univerAPI.createWorkbook({});
  
  // 添加示例数据
  const activeSheet = univerAPI.getActiveWorkbook()?.getActiveSheet();
  if (activeSheet) {
    // 基础数据
    activeSheet.getRange("A1").setValue("数值A");
    activeSheet.getRange("B1").setValue("数值B");
    activeSheet.getRange("C1").setValue("数值C");
    activeSheet.getRange("D1").setValue("结果");
    
    activeSheet.getRange("A2").setValue(2);
    activeSheet.getRange("B2").setValue(3);
    activeSheet.getRange("C2").setValue(5);
    
    activeSheet.getRange("A3").setValue(4);
    activeSheet.getRange("B3").setValue(5);
    activeSheet.getRange("C3").setValue(6);
    
    activeSheet.getRange("A4").setValue(1);
    activeSheet.getRange("B4").setValue(2);
    activeSheet.getRange("C4").setValue(3);
    
    // 使用自定义公式
    // CUSTOM_MULTIPLY_ADD(A2, B2, C2) = 2 * 3 + 5 = 11
    activeSheet.getRange("D2").setFormula("=CUSTOM_MULTIPLY_ADD(A2, B2, C2)");
    
    // CUSTOM_POWER_SUM(A2, B2, C2) = 2² + 3² + 5² = 4 + 9 + 25 = 38
    activeSheet.getRange("D3").setFormula("=CUSTOM_POWER_SUM(A3, B3, C3)");
    
    // CUSTOM_POWER_SUM(A4, B4) = 1² + 2² = 1 + 4 = 5
    activeSheet.getRange("D4").setFormula("=CUSTOM_POWER_SUM(A4, B4)");
    
    // 添加说明
    activeSheet.getRange("F1").setValue("自定义公式说明");
    activeSheet.getRange("F2").setValue("CUSTOM_MULTIPLY_ADD(a,b,c)");
    activeSheet.getRange("G2").setValue("计算 a * b + c");
    activeSheet.getRange("F3").setValue("CUSTOM_POWER_SUM(a,b,...)");
    activeSheet.getRange("G3").setValue("计算所有参数的平方和");
  }

  univerInstance = univer;
  univerAPIInstance = univerAPI;
});

onBeforeUnmount(() => {
  univerInstance?.dispose();
  univerAPIInstance?.dispose();
  univerInstance = null;
  univerAPIInstance = null;
});

// 保存数据到后端
const saveData = async () => {
  if (!univerAPIInstance || !workbook) {
    ElMessage.error("表格未初始化完成");
    return;
  }

  try {
    // 获取当前工作簿数据
    const workbookData = workbook.save() as IWorkbookData;
    
    // 发送数据到后端
    const response:any = await saveSheetData(workbookData);
    
    if (response.code === 200) {
      ElMessage.success("数据保存成功");
    } else {
      ElMessage.error("数据保存失败: " + response.msg);
    }
  } catch (error) {
    console.error("保存数据时出错:", error);
    ElMessage.error("保存数据时出错");
  }
};
</script>

<style scoped>
.univer-sheets-container {
  width: 100%;
  height: 100%;
  padding: 20px;
}

.toolbar {
  margin-bottom: 10px;
}

.save-button {
  padding: 8px 16px;
  background-color: #409eff;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
}

.save-button:hover {
  background-color: #337ecc;
}

.univer-sheet {
  width: 100%;
  height: 800px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;  
}
</style>