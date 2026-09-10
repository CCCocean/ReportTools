
namespace 报表工具
{
    internal interface IProduct
    {
        /// <summary>
        /// 合同编号
        /// </summary>
        string ContractNumber { get; set; }
        /// <summary>
        /// 产品名称
        /// </summary>
        string Name { get; set; }
        /// <summary>
        /// 物料编码
        /// </summary>
        string Code { get; set; }
        /// <summary>
        /// 规格
        /// </summary>
        string Specification { get; set; }
        /// <summary>
        /// 文件目录
        /// </summary>
        string Path { get; set; }
        /// <summary>
        /// 数量
        /// </summary>
        int Quantity { get; set; }
        /// <summary>
        /// 旋转策略
        /// </summary>
        int Rotation { get; set; }

    }
}
