using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gridCopyToClipboardExample {
    class Employee {
        public Employee(int id, string name, double salary, double bonus, string department, DateTime hired) {
            Id = id;
            Name = name;
            Salary = salary;
            Bonus = bonus;
            Department = department;
            Hired = hired;
        }

        public int Id { get; private set; }
        public string Name { get; private set; }
        public double Salary { get; private set; }
        public double Bonus { get; private set; }
        public string Department { get; private set; }
        public DateTime Hired { get; private set; }
    }
}
