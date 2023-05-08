using System;

class Tramas
{
    public static String ArrayToString(UInt16[] dataToClean)
    {
        String stream = "";

        for (int i = 0; i < dataToClean.Length; i++)
        {
            stream = stream + Convert.ToString(dataToClean[i], 16);
        }
        return stream;
    }

    public static String CleanStream(String stringToClean)
    {

        String stream = "";

        for (int i = 4; i <= 14; i++)
        {
            if (i % 2 == 0)
            {
                stream = stream + stringToClean[i];
            }
        }
        return "";
    }
}
