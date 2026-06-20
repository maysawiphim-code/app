import bcrypt

# รหัสผ่านที่ต้องการตั้งสำหรับ user1 (เช่น 1234)
password = "1234" 

# สร้าง Hash
hashed = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

# พิมพ์ออกมาเพื่อเอาไปก๊อปปี้
print(f"เอาค่านี้ไปใส่ในไฟล์ config.yaml:")
print(f'"{hashed.decode("utf-8")}"')