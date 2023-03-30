using System.Drawing;

namespace Soubu.Helper.NPOI
{
    /// <summary>
    /// 单元格属性
    /// </summary>
    public class SoubuNPOICell_M
    {
        private int rowNum;
        private int cellNum;
        private bool isMerge;
        private Point startMerge;
        private Point endMerge;
        private int colspan;
        private int rowspan;
        /// <summary>
        /// 跨列数【即合并的列数】
        /// <para>根据StartMerge和EndMerge计算得出</para>
        /// </summary>
        public int Colspan
        {
            get
            {
                int result = endMerge.Y - startMerge.Y + 1;
                return result < 2 ? 0 : result;
            }
            set { colspan = value; }
        }
        /// <summary>
        /// 跨行数【即合并的行数】
        /// <para>根据StartMerge和EndMerge计算得出</para>
        /// </summary>
        public int Rowspan
        {
            get
            {
                int result = endMerge.X - startMerge.X + 1;
                return result < 2 ? 0 : result;
            }
            set { rowspan = value; }
        }
        private string? content;

        /// <summary>
        /// Excel行号
        /// </summary>
        public int RowNum
        {
            get { return rowNum; }
            set { rowNum = value; }
        }
        /// <summary>
        /// Excel行号对应的列号
        /// </summary>
        public int CellNum
        {
            get { return cellNum; }
            set { cellNum = value; }
        }
        /// <summary>
        /// 当前单元格是否被合并
        /// </summary>
        public bool IsMerge
        {
            get { return isMerge; }
            set { isMerge = value; }
        }
        /// <summary>
        /// 开始合并的单元格坐标
        /// </summary>
        public Point StartMerge
        {
            get { return startMerge; }
            set { startMerge = value; }
        }
        /// <summary>
        /// 结束合并的单元格坐标
        /// </summary>
        public Point EndMerge
        {
            get { return endMerge; }
            set { endMerge = value; }
        }
        /// <summary>
        /// 单元格的内容
        /// </summary>
        public string? Content
        {
            get { return content; }
            set { content = value; }
        }
    }
}