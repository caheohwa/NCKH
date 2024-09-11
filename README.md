# Truy cập thư viện

import pandas as pd
import numpy as np
# Nhập giá trị từ excel
LDT = pd.read_excel('D:/NCKH/linedata.xlsx')
BDT = pd.read_excel('D:/NCKH/busdata.xlsx')

# Lưu các giá trị từ excel
typebus = BDT['loainut']
chieudai = LDT['l'] * 10**-3
r = LDT['r'] * chieudai
x = LDT['x'] * chieudai
n = BDT.shape[0]
nh = LDT.shape[0]
rb = BDT['rb']
xb = BDT['xb']
tsobien = BDT['k']
Sphutai = BDT['s']
cosphi = BDT['cosphi']
dp0 = BDT['dp0'] * 10**-3
dq0 = BDT['dq0'] * 10**-3

# Khởi tạo dữ liệu ban đầu
Scb = 100              #Đơn vị MWA
Ucb = BDT.iloc[0, 3]   #Lấy bằng Uđm, đơn vị KV
DC = 100               #Giá trị lặp ban đầu
esp = 0.001            #Sai số so sánh
# Tính toán các giá trị
U = BDT['u']/Ucb
zb = rb + 1j*xb
P = Sphutai * cosphi
# Khởi tạo danh sách rỗng để lưu kết quả Q
Q_list = []
# Sử dụng vòng lặp for để tính toán Q
for i in range(n):
    if cosphi[i] == 0:
        Q = Sphutai[i]
    else:
        Q = np.sqrt(Sphutai[i]**2 - P[i]**2)
    Q_list.append(Q)

# Chuyển danh sách Q_list thành Series hoặc cột trong DataFrame
Q = pd.Series(Q_list)
P = P * 10**-3
Q = Q * 10**-3
DSb = (P**2 + Q**2)/Ucb**2 * zb
# Chuyển DSb từ Series pandas sang mảng numpy để lấy phần thực và phần ảo
DSb_real = np.real(DSb.values)
DSb_imag = np.imag(DSb.values)
P = P + DSb_real + dp0
Q = Q + DSb_imag + dq0
P = P/Scb
Q = Q/Scb
Icb = Scb/(np.sqrt(3)*Ucb)
Zcb = Ucb**2/Scb
buoclap = 0
# Chương trình con Ybus
# Khởi tạo ma trận Ybus với kích thước n x n và giá trị ban đầu là 0 + 0j
Ybus = np.zeros((n, n), dtype=complex)
# Tính toán phần tử ngoài đường chéo chính của Ybus
for e in range(nh):
    nut_dau = LDT.iloc[e, 1] - 1  # Chỉ số Python bắt đầu từ 0
    nut_cuoi = LDT.iloc[e, 2] - 1
    Ybus[nut_dau, nut_cuoi] = -1 / (r[e] + 1j * x[e]) * Zcb
    Ybus[nut_cuoi, nut_dau] = Ybus[nut_dau, nut_cuoi]  # Đối xứng

# Tính toán phần tử trên đường chéo chính của Ybus
for e in range(n):
    Ybus[e, e] = 0  # Khởi tạo phần tử trên đường chéo chính
    for f in range(n):
        if f != e:
            Ybus[e, e] -= Ybus[e, f]
# Giả sử Ybus đã được tính toán từ đoạn mã trước đó
Ybus_df = pd.DataFrame(Ybus, index=[f'Nút {i+1}' for i in range(n)], columns=[f'Nút {i+1}' for i in range(n)])

# Xuất ma trận Ybus ra file Excel
Ybus_df.to_excel('Ybus_output.xlsx', index=True)
print("Ma trận Ybus đã được xuất ra file 'Ybus_output.xlsx'.")
# Lấy giá trị tuyệt đối của các phần tử trong Ybus
Y_abs = np.abs(Ybus)

# Tìm góc lệch pha (góc của số phức) của các phần tử trong Ybus
angles = np.angle(Ybus)

# Tạo ma trận d (vector cột với n phần tử, tất cả đều bằng 0)
d = np.zeros((n, 1))

# Chuyển Y_abs và angles thành DataFrame
Y_abs_df = pd.DataFrame(Y_abs, index=[f'Nút {i+1}' for i in range(n)], columns=[f'Nút {i+1}' for i in range(n)])
angles_df = pd.DataFrame(angles, index=[f'Nút {i+1}' for i in range(n)], columns=[f'Nút {i+1}' for i in range(n)])

# Tạo một file Excel với các sheet khác nhau cho Y_abs và angles
with pd.ExcelWriter('Ybus_results.xlsx') as writer:
    Y_abs_df.to_excel(writer, sheet_name='Y_abs')
    angles_df.to_excel(writer, sheet_name='Angles')

print("Ma trận Y_abs và góc lệch pha đã được xuất ra file 'Ybus_results.xlsx'.")
### Chương trình con công suất cho nhận

m = 0
for e in range(1, n):
    if typebus[e] == 2: #Xác nhận nút PV
        m += 1

Pcho = [0] * (n - 1)
Qcho = [0] * (n - 1 - m)
f, g = 0, 0

for e in range(2, n):
    if typebus[e] == 2:
        Pcho[f] = P[e]
        f += 1
    elif typebus[e] == 3: #Xác nhận nút PQ
        Pcho[f] = -P[e]
        Qcho[g] = -Q[e]
        f += 1
        g += 1
# MAIN PROGRAM#
while DC>esp:
    buoclap +=1
# CHƯƠNG TRÌNH CON TÍNH ĐỘ LỆCH CÔNG SUẤT
# Bên trong vòng lặp while
# Tạo mảng P1 với kích thước n x 1
    P1 = np.zeros((n, 1))
# Tính ràng buộc công suất Pi
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        P1[e] = 0
        for f in range(n):
            P1[e] += U[e] * U[f] * Ybus[e, f] * np.cos(angles[e, f] - d[e] + d[f])
 # Bỏ phần tử P1[0] (tương đương P1(1) = [] trong MATLAB)
    P1 = P1[1:]  # Vì P1[1] bỏ qua nút nguồn, giữ lại từ phần tử 2 trở đi
# Tạo mảng Q1 với kích thước (n-1-m) x 1
    Q1 = np.zeros((n - 1 - m, 1))
    g = 0
# Tính ràng buộc công suất Qi
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        if typebus[e] == 3:  # Xác nhận nút PQ
            Q2 = 0
            for f in range(n):
                Q2 -= U[e] * U[f] * Ybus[e, f] * np.sin(angles[e, f] - d[e] + d[f])
            Q1[g] = Q2
            g += 1
# Tính độ lệch công suất
    DP = Pcho - P1.flatten()  # Chuyển mảng 2D thành 1D để trừ
    DQ = Qcho - Q1.flatten()
# Ghép DP và DQ thành DC1
    DC1 = np.concatenate((DP, DQ), axis=0)
# MA TRẬN JACOBI#
 # Tạo ma trận J1 với kích thước n x n
    J1 = np.zeros((n, n))
# Phần tử ngoài đường chéo
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        for f in range(2, n):
            if f != e:
                J1[e, f] = -U[e] * U[f] * Ybus[e, f] * np.sin(angles[e, f] - d[e] + d[f])
# Phần tử trên đường chéo
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        J1[e, e] = 0
        for f in range(n):
            if f != e:
                J1[e, e] += U[e] * U[f] * Ybus[e, f] * np.sin(angles[e, f] - d[e] + d[f])
# Tạo ma trận J2 với kích thước n x n
    J2 = np.zeros((n, n))
 # Phần tử ngoài đường chéo
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        for f in range(2, n):
            if f != e and typebus[f] == 3:  # Kiểm tra điều kiện ngoại trừ đường chéo và typebus[f] == 3
                J2[e, f] = U[e] * U[f] * Ybus[e, f] * np.cos(angles[e, f] - d[e] + d[f])
# Phần tử trên đường chéo
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        if typebus[e] == 3:  # Kiểm tra typebus[e] == 3
            J2[e, e] = 0
            for f in range(n):
                if f != e:
                    J2[e, e] += U[e] * U[f] * Ybus[e, f] * np.cos(angles[e, f] - d[e] + d[f])
# Cộng thêm phần tử trên đường chéo
            J2[e, e] += 2 * U[e] * U[e] * Ybus[e, e] * np.cos(angles[e, e])
# Tạo ma trận J3 với kích thước n x n
    J3 = np.zeros((n, n))
 # Phần tử ngoài đường chéo
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        if typebus[e] == 3:  # Kiểm tra điều kiện typebus[e] == 3
            for f in range(2, n):
                if f != e:  # Ngoại trừ đường chéo
                    J3[e, f] = -U[e] * U[f] * Ybus[e, f] * np.cos(angles[e, f] - d[e] + d[f])
# Phần tử trên đường chéo
    for e in range(2, n):  # Bắt đầu từ 2, bỏ qua nút nguồn
        if typebus[e] == 3:  # Kiểm tra điều kiện typebus[e] == 3
            J3[e, e] = 0
            for f in range(1, n):
                if f != e:
                    J3[e, e] += U[e] * U[f] * Ybus[e, f] * np.cos(angles[e, f] - d[e] + d[f])




