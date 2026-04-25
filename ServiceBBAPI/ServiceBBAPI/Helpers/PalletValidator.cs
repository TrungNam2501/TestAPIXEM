using System.Collections.Generic;

namespace ServiceBBAPI.Helpers
{
    public static class PalletValidator
    {
        private static readonly HashSet<string> ValidPrefixes = new HashSet<string>
        {
            "VB", "EB", "VC", "VD", "VE", "VF"
        };

        public static string Validate(string pallet)
        {
            if (string.IsNullOrWhiteSpace(pallet))
                return "Pallet không được để trống!";

            pallet = pallet.Trim();

            if (pallet.Length < 6)
                return "Pallet không đủ ký tự, nhập lại!!!";

            string prefix2 = pallet.Substring(0, 2);
            string prefix1 = pallet.Substring(0, 1);

            if (!ValidPrefixes.Contains(prefix2) && prefix1 != "V")
                return "Pallet Không hợp lệ, nhập lại!!!";

            return null;
        }
    }
}
