using System;
using System.Linq;
using System.Collections.Generic;
	
namespace MvcHomeWork.Models
{   
	public  class V_客戶資料清單Repository : EFRepository<V_客戶資料清單>, IV_客戶資料清單Repository
	{

	}

	public  interface IV_客戶資料清單Repository : IRepository<V_客戶資料清單>
	{

	}
}