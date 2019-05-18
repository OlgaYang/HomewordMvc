using System;
using System.ComponentModel.DataAnnotations;

namespace MvcHomeWork.DataTypeAttributes
{
    public class CellPhoneFormatValidateAttribute : DataTypeAttribute
    {
        public CellPhoneFormatValidateAttribute() : base(DataType.Text)
        {
        }

        public override bool IsValid(object value)
        {
            if (value == null)
                return true;

            string phoneNumber = (string)value;

            if (phoneNumber.Length != 11)
            {
                ErrorMessage = "手機格式必須為11碼";
                return false;
            }

            for (int i = 0; i < phoneNumber.Length; i++)
            {
                if (i < 4 || i > 4)
                {
                    if (!Char.IsNumber(phoneNumber[i]))
                    {
                        ErrorMessage = "手機前四碼跟後六碼，只能為數字";
                        return false;
                    }
                }
                else
                {
                    if (phoneNumber[i] != '-')
                    {
                        ErrorMessage = "手機中間必須由-隔開";
                        return false;
                    }
                }
            }

            return true;
        }
    }
}