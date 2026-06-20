import bcrypt

passwords = ['1234', 'adminpassword']

for pwd in passwords:
    # แปลงรหัสผ่านเป็น bytes และ hash
    hashed = bcrypt.hashpw(pwd.encode('utf-8'), bcrypt.gensalt())
    # แสดงผลในรูปแบบ string เพื่อนำไปใส่ใน config.yaml
    print(f"Password: {pwd} -> Hash: {hashed.decode('utf-8')}")