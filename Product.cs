using System.IO;

namespace 报表工具
{

    public class Product : IProduct
    {
        // 密度常量
        private const double Density = 7.85;

        private string _contractNumber;
        private string _name;
        private string _code;
        private string _specification;
        private string _path;
        private int _quantity;
        private int _rotation;
        private double _length;
        private double _width;
        private double _area;
        private double _thickness;
        private int _toPackQty;

        public string ContractNumber
        {
            get => _contractNumber;
            set => _contractNumber = value;
        }

        public string Name
        {
            get => _name;
            set => _name = value;
        }

        public string Code
        {
            get => _code;
            set => _code = value;
        }

        public string Specification
        {
            get => _specification;
            set => _specification = value;
        }

        public string Path
        {
            get => _path;
            set => _path = value;
        }

        public int Quantity
        {
            get => _quantity;
            set => _quantity = value;
        }

        public int Rotation
        {
            get => _rotation;
            set => _rotation = value;
        }

        public double Length
        {
            get => _length;
            set => _length = value;
        }

        public double Width
        {
            get => _width;
            set => _width = value;
        }

        public double Area
        {
            get => _area;
            set => _area = value;
        }

        public double Thickness
        {
            get => _thickness;
            set => _thickness = value;
        }

        public int ToPackQty
        {
            get => _toPackQty;
            set => _toPackQty = value;
        }

        public double Weight
        {
            get => Area * Thickness * Density;
        }

        /// <summary>
        /// 验证文件是否存在
        /// </summary>
        /// <returns>如果Path为null或空，返回false；否则返回文件是否存在</returns>
        public bool PathExists()
        {
            if (string.IsNullOrWhiteSpace(_path))
                return false;
            return File.Exists(_path);
        }
    }
}
