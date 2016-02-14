using Microsoft.Lync.Model.Conversation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace BusinessCats
{
    public class LyncParticipant
    {
        public Participant participant;
        public InstantMessageModality modality;
        public ECDiffieHellmanCng dh;
        public byte[] derivedKey;

        public byte[] GetPublicKey()
        {
            if (dh == null)
            {
                dh = new ECDiffieHellmanCng() { KeyDerivationFunction = ECDiffieHellmanKeyDerivationFunction.Hash, HashAlgorithm = CngAlgorithm.Sha256 };
            }

            return dh.PublicKey.ToByteArray();
        }

        public void DeriveKey(byte[] keyBlob)
        {
            var otherKey = CngKey.Import(keyBlob, CngKeyBlobFormat.EccPublicBlob);
            derivedKey = dh.DeriveKeyMaterial(otherKey);
        }
    }
}
