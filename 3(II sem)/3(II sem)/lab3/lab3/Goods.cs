using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab3
{
    public class Goods
    {
        public int id;
        public string product;
        public int count;
        public int price;
        public int sum;

        public int Id { get; set; }
        public string Product { get; set; }
        public int Count { get; set; }
        public int Price { get; set; }

        public int Sum
        {
            get
            {
                return Price * Count;
            }
            set
            {
                value = Price * Count;
            }
        }

        public Goods() { }

        public List<Goods> Initialize()
        {
            return new List<Goods>()
            {
              new Goods
              {
                Id=1,
                Product="Ананасы",
                Count=3,
                Price=16,
                Sum = Price * Count,
              },
              new Goods
              {
                Id=2,
                Product="Апельсины",
                Count=10,
                Price=40,
                Sum=Price * Count,
              },
              new Goods
              {
                Id=3,
                Product="Яблоки",
                Count=8,
                Price=80,
                Sum = Price * Count,
              },
              new Goods
              {
                Id=4,
                Product="Лимоны",
                Count=35,
                Price=40,
                Sum = Price * Count,
              },
           };
        }
    }
}
