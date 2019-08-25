using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace License
{
    public static class LicenseHandler
    {
        public static string GenerateUid(string appName)
        {
            return HardwareInfo.GenerateUid(appName);
        }

        private static LicenseEntity ReadLicense(Type licenseObjType, string licenseString, byte[] certPubKeyData, out LicenseStatus licStatus, out string validationMsg)
        {
            if (string.IsNullOrWhiteSpace(licenseString))
            {
                licStatus = LicenseStatus.Cracked;
                validationMsg = "Licencja uszkodzona";
                return null;
            }

            LicenseEntity license = null;

            try
            {
                //Get RSA key from certificate
                X509Certificate2 cert = new X509Certificate2(certPubKeyData);
                RSACryptoServiceProvider rsaKey = (RSACryptoServiceProvider)cert.PublicKey.Key;

                XmlDocument xmlDoc = new XmlDocument {PreserveWhitespace = true};

                // Load an XML file into the XmlDocument object.
                xmlDoc.LoadXml(Encoding.UTF8.GetString(Convert.FromBase64String(licenseString)));

                // Verify the signature of the signed XML.            
                if (VerifyXml(xmlDoc, rsaKey))
                {
                    XmlNodeList nodeList = xmlDoc.GetElementsByTagName("Signature");
                    xmlDoc.DocumentElement?.RemoveChild(nodeList[0]);

                    string licXml = xmlDoc.OuterXml;

                    //Deserialize license
                    XmlSerializer serializer = new XmlSerializer(typeof(LicenseEntity), new[] { licenseObjType });

                    using (StringReader reader = new StringReader(licXml))
                    {
                        license = (LicenseEntity)serializer.Deserialize(reader);
                    }

                    licStatus = license.DoExtraValidation(out validationMsg);
                }
                else
                {
                    licStatus = LicenseStatus.Invalid;
                    validationMsg = "Nieprawidłowy plik licencji";
                }
            }
            catch
            {
                licStatus = LicenseStatus.Cracked;
                validationMsg = "Licencja uszkodzona";
            }

            return license;
        }
        
        // Verify the signature of an XML file against an asymmetric 
        // algorithm and return the result.
        private static bool VerifyXml(XmlDocument doc, RSA key)
        {
            // Check arguments.
            if (doc == null)
                throw new ArgumentException("Doc");

            if (key == null)
                throw new ArgumentException("Key");

            // Create a new SignedXml object and pass it
            // the XML document class.
            SignedXml signedXml = new SignedXml(doc);

            // Find the "Signature" node and create a new
            // XmlNodeList object.
            XmlNodeList nodeList = doc.GetElementsByTagName("Signature");

            // Throw an exception if no signature was found.
            if (nodeList.Count <= 0)
            {
                throw new CryptographicException("Verification failed: No Signature was found in the document.");
            }

            // This example only supports one signature for
            // the entire XML document.  Throw an exception 
            // if more than one signature was found.
            if (nodeList.Count >= 2)
            {
                throw new CryptographicException("Verification failed: More that one signature was found for the document.");
            }

            // Load the first <signature> node.  
            signedXml.LoadXml((XmlElement)nodeList[0]);

            // Check the signature and return the result.
            return signedXml.CheckSignature(key);
        }
        
        public static MyLicense ReadLicense(out LicenseStatus licStatus, out string validationMsg)
        {
            byte[] certPublicKeyData;

            Assembly assembly = Assembly.GetExecutingAssembly();

            using (MemoryStream memoryStream = new MemoryStream())
            {
                assembly.GetManifestResourceStream("ScanHelper.LicenseVerify.cer")?.CopyTo(memoryStream);

                certPublicKeyData = memoryStream.ToArray();
            }

            if (File.Exists("license.lic"))
            {
                MyLicense license = (MyLicense)ReadLicense(typeof(MyLicense), File.ReadAllText("license.lic"), certPublicKeyData, out licStatus, out validationMsg);

                return license;
            }

            licStatus = LicenseStatus.Undefined;
            validationMsg = "Brak pliku licencji!";

            return null;
        }
    }

}
