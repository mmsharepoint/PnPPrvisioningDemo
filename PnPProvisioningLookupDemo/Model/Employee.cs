using System;
using System.Collections.Generic;
using System.Text;

namespace PnPProvisioningLookupDemo.Model
{
  class Employee
  {
    public string Title { get; set; }
    public string Firstname { get; set; }
    public string Lastname { get; set; }
    public string Street { get; set; }
    public string StreetNo { get; set; }
    public int EmployeeLocation { get; set; }
    public int OldEmployeeLocation { get; set; }
    public string Salary { get; set; }
  }
}
