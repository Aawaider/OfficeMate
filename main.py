import ssl
import datetime
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa, ec
from cryptography.hazmat.backends import default_backend
import os

def generate_enhanced_self_signed_cert(cert_file="certificate.pem", 
                                     key_file="private.key",
                                     common_name="localhost",
                                     days_valid=365,
                                     key_type="rsa",
                                     key_size=2048):
    """
    增强版自签名证书生成
    
    Args:
        key_type: "rsa" 或 "ec" (椭圆曲线)
        key_size: RSA密钥长度或EC曲线类型
    """
    
    # 生成私钥
    if key_type.lower() == "ec":
        # 使用椭圆曲线密码学
        if key_size == 256:
            curve = ec.SECP256R1()
        elif key_size == 384:
            curve = ec.SECP384R1()
        else:
            curve = ec.SECP256R1()
        
        private_key = ec.generate_private_key(curve, default_backend())
    else:
        # 使用RSA
        private_key = rsa.generate_private_key(
            public_exponent=65537,
            key_size=key_size,
            backend=default_backend()
        )
    
    # 创建证书主题
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "CN"),
        x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME, "Beijing"),
        x509.NameAttribute(NameOID.LOCALITY_NAME, "Beijing"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "My Organization"),
        x509.NameAttribute(NameOID.ORGANIZATIONAL_UNIT_NAME, "IT Department"),
        x509.NameAttribute(NameOID.COMMON_NAME, common_name),
    ])
    
    # 修复：使用时区感知的时间
    current_time = datetime.datetime.now(datetime.timezone.utc)
    not_valid_after = current_time + datetime.timedelta(days=days_valid)
    
    # 构建证书
    builder = x509.CertificateBuilder()
    builder = builder.subject_name(subject)
    builder = builder.issuer_name(issuer)
    builder = builder.public_key(private_key.public_key())
    builder = builder.serial_number(x509.random_serial_number())
    builder = builder.not_valid_before(current_time)
    builder = builder.not_valid_after(not_valid_after)
    
    # 修复：正确的IP地址类名
    san_list = [
        x509.DNSName(common_name),
        x509.DNSName(f"*.{common_name}"),
        x509.DNSName("localhost"),
    ]
    
    # 添加IP地址（可选）
    try:
        san_list.append(x509.IPAddress(x509.IPv4Address("127.0.0.1")))
    except Exception as e:
        print(f"警告: 无法添加IP地址到SAN扩展: {e}")
    
    # 添加扩展
    builder = builder.add_extension(
        x509.SubjectAlternativeName(san_list),
        critical=False,
    )
    
    # 添加密钥用途扩展
    builder = builder.add_extension(
        x509.KeyUsage(
            digital_signature=True,
            key_encipherment=True,
            key_cert_sign=False,
            crl_sign=False,
            content_commitment=False,
            data_encipherment=False,
            key_agreement=False,
            encipher_only=False,
            decipher_only=False
        ),
        critical=True
    )
    
    # 添加扩展密钥用途
    builder = builder.add_extension(
        x509.ExtendedKeyUsage([
            x509.oid.ExtendedKeyUsageOID.SERVER_AUTH,
            x509.oid.ExtendedKeyUsageOID.CLIENT_AUTH,
        ]),
        critical=False
    )
    
    # 添加基本约束
    builder = builder.add_extension(
        x509.BasicConstraints(ca=False, path_length=None),
        critical=True
    )
    
    # 签名证书
    certificate = builder.sign(
        private_key=private_key,
        algorithm=hashes.SHA256(),
        backend=default_backend()
    )
    
    # 保存文件
    os.makedirs(os.path.dirname(cert_file) if os.path.dirname(cert_file) else '.', exist_ok=True)
    
    with open(cert_file, "wb") as f:
        f.write(certificate.public_bytes(serialization.Encoding.PEM))
    
    with open(key_file, "wb") as f:
        f.write(private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption()
        ))
    
    # 设置文件权限（仅限Unix系统）
    try:
        os.chmod(key_file, 0o600)
    except:
        pass
    
    print(f"✅ 证书已生成: {cert_file}")
    print(f"✅ 私钥已生成: {key_file}")
    print(f"📅 有效期: {days_valid}天")
    print(f"🔑 密钥类型: {key_type.upper()}")
    print(f"🌐 通用名称: {common_name}")
    
    return cert_file, key_file

# 使用示例
if __name__ == "__main__":
    # 生成RSA证书
    cert, key = generate_enhanced_self_signed_cert(
        cert_file="my_cert.pem",
        key_file="my_key.pem",
        common_name="mysite.example.com",
        days_valid=365,
        key_type="rsa",
        key_size=2048
    )