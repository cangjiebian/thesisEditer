using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentSearch
{
    class Book
    {
        public string Title;
        public string Author;
        public string Place;
        public string Press;
        public string Date;
        public string Page;
        public string Price;
        public string CallNumber;
        public string ISBN;
        public string Link;
        public Book(string title, string author, string place, string press, string date, string page, string price,
            string callnumber, string isbn, string link)
        {
            Title = title;
            Author = author;
            Place = place;
            Press = press;
            Date = date;
            Page = page;
            Price = price;
            CallNumber = callnumber;
            ISBN = isbn;
            Link = link;
        }



    }
}
