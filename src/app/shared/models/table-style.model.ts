export interface TableStyleConfig {
  style: string; // 表格风格
  titleColor: string; // 标题栏背景颜色
  headerColor: string; // 表头背景颜色
  totalColor: string; // 合计背景颜色
  fontFamily: string; // 字体选择
  titleFontSize: number; // 标题字体大小
  titleFontBold: boolean; // 标题是否加粗
  titleFontColor: string; // 标题字体颜色
  headerFontSize: number; // 表头字体大小
  headerFontBold: boolean; // 表头是否加粗
  headerFontColor: string; // 表头字体颜色
  dataFontSize: number; // 数据字体大小
  dataFontBold: boolean; // 数据是否加粗
  dataFontColor: string; // 数据字体颜色
  borderColor: string; // 边框颜色
  borderStyle: string; // 边框样式
}

export const DEFAULT_TABLE_STYLE: TableStyleConfig = {
  style: '商务风格',
  titleColor: '#4472C4',
  headerColor: '#5B9BD5',
  totalColor: '#F2F2F2',
  fontFamily: '微软雅黑',
  titleFontSize: 14,
  titleFontBold: true,
  titleFontColor: '#FFFFFF',
  headerFontSize: 11,
  headerFontBold: true,
  headerFontColor: '#FFFFFF',
  dataFontSize: 10,
  dataFontBold: false,
  dataFontColor: '#000000',
  borderColor: '#CCCCCC',
  borderStyle: 'thin'
};

