export class GlobalProps {
  slideWidth?: number
  slideHeight?: number
  slidePaths?: any[]
  slideLayouts?: any[]
  thumbnail?: string
  theme?: any
}

export class SingleSlide {
  content: any
  // slide的资源内容
  resContent: any
  // 布局文件的内容
  layoutContent: any
  // 布局文件的资源内容
  layoutResContent: any
  // 布局文件内容索引表
  layoutIndexTables?: { idTable: any, idxTable: any, typeTable: any }
  // 母版的内容
  masterContent: any
  // 母版的字体样式
  masterTextStyles?: string
  // 母版的索引表
  masterIndexTable?: { idTable: any, idxTable: any, typeTable: any }
  masterResContent?: any
  bgColor?: string
  gprops?: GlobalProps
}
