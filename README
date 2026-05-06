# 💉 DeltaMV — Vietnam Vaccination Behavior Analysis

> **Phân tích hành vi tiêm chủng vaccine tại thị trường Việt Nam**  
> *Dựa trên dữ liệu khảo sát tại phòng khám VNVC & Long Châu | M1 – M3/2025*


## 🛠️ Tools & Technologies

| Tool | Purpose |
|------|---------|
| **Power BI** | Dashboard design & interactive reporting |
| **Microsoft Excel** | Data staging, BRD documentation, survey mapping |
| **DAX** | KPI calculation & custom measures |
| **Survey / Exit Interview Data** | Primary data source (collected at clinic point-of-care) |

---

## 1. Background and Overview

### Project Background

Dự án DeltaMV được thực hiện nhằm hỗ trợ đối tác dược phẩm (GSK và các hãng vaccine lớn) đánh giá hành vi tiêm chủng thực tế của người dân tại hai kênh phân phối chính tại Việt Nam: **VNVC** và **Long Châu**. Dữ liệu được thu thập thông qua khảo sát exit interview trực tiếp tại phòng khám, ghi nhận cả ý định trước và hành vi thực tế sau khi khách hàng tiêm chủng.

Nhóm khách hàng trọng tâm được phân tích là **người lớn từ 50 tuổi trở lên** — phân khúc có nhu cầu tiêm chủng cao nhất đối với các vaccine như Shingles (Zona thần kinh), Pneumococcal (Phế cầu khuẩn), và Influenza (Cúm) — cùng với nhóm **trẻ em từ 0–2 tuổi** cho phân tích vaccine Nhi khoa.

### Business Question

> *"Khoảng cách giữa ý định tiêm và hành vi tiêm thực tế là gì? Đâu là các yếu tố ảnh hưởng đến quyết định tiêm chủng và lựa chọn thương hiệu vaccine của người dân Việt Nam?"*

Để trả lời câu hỏi này, phân tích được triển khai theo 4 hướng chính:

- 💊 **Vaccine Units** — Số lượng và cơ cấu vaccine tiêu thụ theo thời gian, khu vực, kênh phân phối
- 📊 **Vaccination Share** — Thị phần các hãng dược phẩm theo từng nhóm bệnh
- 🏥 **Intention to Clinic** — Tỷ lệ intention vs. actual và lý do đến tiêm
- 🎯 **Shingles Intention** — Phân tích chuyên sâu hành vi tiêm Zona thần kinh và vai trò của HCP
- 📣 **Clinic Introduction** — Hiệu quả khuyến nghị của nhân viên y tế đối với nhóm không có ý định ban đầu
- 📢 **Clinic Communication** — Nhận biết thương hiệu và hiệu quả vật liệu truyền thông tại phòng khám

### Project Objectives

- Xây dựng hệ thống báo cáo Power BI đa chiều để theo dõi hành vi tiêm chủng theo thời gian và địa lý
- Đo lường khoảng cách intention–actual gap theo từng loại vaccine và nhóm bệnh
- Đánh giá hiệu quả của HCP (Healthcare Professional) trong việc khuyến nghị và chuyển đổi hành vi tiêm
- Phân tích thị phần hãng dược phẩm và lý do lựa chọn thương hiệu vaccine
- Cung cấp insight hỗ trợ quyết định marketing và chiến lược triển khai dịch vụ tiêm chủng

**Phạm vi dữ liệu:** M1 – M3/2025 | Khảo sát exit interview tại phòng khám | Tập trung: nhóm 50+ và nhóm Nhi (0–2 tuổi)

---

## 2. Data Structure Overview

Dữ liệu được tổ chức theo mô hình **Star Schema** với 1 nhóm bảng Fact và nhiều bảng Dimension, staging trong file `STG_DeltaMV.xlsx`. Dưới đây là tổng quan các bảng dữ liệu:

### Dimension Tables

| Bảng | Mô tả |
|------|-------|
| `Dim_AgeGroup` | Nhóm tuổi khách hàng (0-2, 2-5, 5-10, 10-18, 18-24, 25-32, 33-40, 41-49, 50+) |
| `Dim_Brand` | Danh sách hãng dược phẩm (Sanofi, GSK, Pfizer, MSD, Abbott, Takeda…) |
| `Dim_Center` | Phòng khám và kênh phân phối (VNVC / Long Châu × Hà Nội / HCM) |
| `Dim_Customer` | Hồ sơ khách hàng (tên, tuổi, giới tính, loại: Adult/Child) |
| `Dim_Disease` | Nhóm bệnh được tiêm vaccine (Shingles, Influenza, Pneumococcal, Dengue, Hepatitis…) |
| `Dim_Dose` | Loại mũi tiêm (Initial dose, Booster dose) |
| `Dim_Doctor` | Loại chuyên gia y tế khuyến nghị |
| `Dim_HealthCondition` | Bệnh lý nền của khách hàng |
| `Dim_IntentionReason` | Lý do có ý định tiêm (bác sĩ khuyến nghị, tự tìm hiểu, người thân gợi ý…) |
| `Dim_Suggestor` | Người khuyến nghị tại phòng khám |
| `Dim_Version` | Đợt khảo sát (M1, M2, M3) |
| `Dim_Region` | Khu vực địa lý (Hà Nội, HCM) |

### Fact Tables

| Bảng | Mô tả |
|------|-------|
| `Fact_Actual` | Hành vi tiêm thực tế: vaccine đã tiêm, liều, thời gian giữa các mũi |
| `Fact_Intention` | Ý định tiêm trước khi đến phòng khám theo từng nhóm bệnh |
| `Fact_BrandIntention` | Ý định chọn thương hiệu vaccine cụ thể |
| `Fact_ActualBrandChoice` | Thương hiệu vaccine thực tế được tiêm và lý do chọn |
| `Fact_ActualDifferent` | Lý do khách hàng tiêm vaccine khác so với ý định ban đầu |
| `Fact_NoIntendButActualGet` | Khách hàng không có ý định nhưng thực tế đã tiêm |
| `Fact_RecSuggestor` | Lịch sử khuyến nghị tại phòng khám theo nhóm bệnh |
| `Fact_RecReasonNotGet` | Lý do từ chối khuyến nghị tiêm vaccine |
| `Fact_Advertisement` | Nhận biết vật liệu truyền thông tại phòng khám |
| `STG HCP RCM(doctor)` | Cuộc trò chuyện giữa khách hàng và HCP (tình huống, người khởi xướng) |
| `STG HCP RCM(health condition)` | Bệnh lý nền được ghi nhận trong cuộc tư vấn HCP |

### Key DAX Measures

```dax
-- Tổng số vaccine tiêu thụ
Total Vaccines = COUNT(Fact_Actual[ActualID])

-- Tỷ lệ mũi tiêm theo nhóm bệnh
% by Disease = 
    DIVIDE(
        COUNT(Fact_Actual[ActualID]),
        CALCULATE(COUNT(Fact_Actual[ActualID]), ALL(Dim_VaccineActualGet[Disease]))
    )

-- Tỷ lệ ý định tiêm (Intention Rate)
% Intention = 
    DIVIDE(
        COUNT(Fact_Reason_Intend[CustomerID]),
        COUNT(Dim_Customer[CustomerID])
    )

-- Tỷ lệ tiêm thực tế (Actual Rate)
% Actual = 
    DIVIDE(
        COUNT(Fact_Actual[CustomerID]),
        COUNT(Dim_Customer[CustomerID])
    )

-- Tỷ lệ chuyển đổi từ khuyến nghị (HCP Conversion)
% No Intend But Actual Get = 
    DIVIDE(
        COUNT(Fact_NoIntendButActualGet[CustomerID]),
        COUNT(Fact_RecSuggestor[CustomerID])
    )
```

---

## 3. Executive Overview

### 🏥 Tổng quan Khảo sát tại Phòng khám (M1–M3/2025)

| Chỉ số | Giá trị |
|--------|---------|
| Tổng số khách hàng khảo sát | **~500+ người** |
| Kênh phân phối | **VNVC & Long Châu** |
| Khu vực | **Hà Nội & TP. Hồ Chí Minh** |
| Nhóm tuổi trọng tâm | **50+ tuổi (Adult) & 0–2 tuổi (Pediatric)** |
| Số đợt khảo sát | **3 đợt (M1, M2, M3)** |
| Số hãng dược phẩm theo dõi | **18 hãng** |
| Số nhóm bệnh phân tích | **10+ nhóm bệnh** |

---

## 4. Insight Deep Dive

### 4.1 Vaccine Units — Tiêu thụ theo nhóm bệnh

**Các nhóm vaccine trọng tâm được theo dõi:**

| Vaccine | Nhóm bệnh | Đối tượng chính |
|---------|-----------|----------------|
| Shingles (Zona thần kinh) | Shingles | 50+ tuổi |
| Pneumococcal (Phế cầu khuẩn) | Pneumococcal | 50+ tuổi |
| Influenza (Cúm) | Influenza | 50+ tuổi & toàn dân |
| Dengue Fever (Sốt xuất huyết) | Dengue | 18–40 tuổi |
| Hepatitis (Viêm gan) | Hepatitis B/A | Đa nhóm tuổi |
| MMR (Sởi, Quai bị, Rubella) | MMR | Trẻ em & phụ nữ |
| Chickenpox (Thủy đậu) | Chickenpox | Trẻ em |
| 6-in-1 | Multi-disease | 0–2 tuổi |

**Findings:**
- Tỷ lệ tiêu thụ vaccine **Shingles và Pneumococcal** có xu hướng tăng qua các đợt khảo sát (M1→M3), phản ánh hiệu quả của các chương trình nâng cao nhận thức
- **Influenza** là vaccine có số lượng tiêu thụ lớn nhất và ổn định nhất theo thời gian
- Sự chênh lệch giữa **tỷ lệ ý định và tỷ lệ thực tế** cao nhất ở vaccine Shingles và Dengue — là khoảng trống cần được tối ưu hóa qua truyền thông và tư vấn tại điểm bán

---

### 4.2 Vaccination Share — Thị phần hãng dược

**Các hãng dược phẩm được theo dõi:**

| Hãng | Vaccine Đại diện |
|------|----------------|
| **GSK** | Shingrix (Shingles), Cervarix (HPV) |
| **Pfizer** | Prevenar 13/20 (Pneumococcal) |
| **MSD** | Gardasil (HPV), MMR II |
| **Sanofi** | Dengvaxia (Dengue), Vaxigrip (Influenza) |
| **Abbott** | Influvac Tetra (Influenza) |
| **Takeda** | Qdenga (Dengue) |

**Findings:**
- **Pfizer (Prevenar 13)** dẫn đầu thị phần ở nhóm Pneumococcal — được bác sĩ khuyến nghị cao hơn đáng kể so với các thương hiệu cạnh tranh
- Lý do lựa chọn thương hiệu chủ đạo là **"bác sĩ khuyến nghị thương hiệu này"** — cho thấy HCP đóng vai trò quyết định trong conversion
- Phần lớn trường hợp **thay đổi thương hiệu so với ý định ban đầu** là do hết hàng tại phòng khám, không phải do thiếu nhận thức

---

### 4.3 Intention to Clinic — Khoảng cách Ý định vs. Thực tế

**Tỷ lệ Intention vs. Actual theo nhóm bệnh (minh họa):**

| Nhóm vaccine | % Intention | % Actual | Gap |
|-------------|:-----------:|:--------:|:---:|
| Shingles | ~35% | ~25% | -10% |
| Pneumococcal | ~40% | ~32% | -8% |
| Influenza | ~55% | ~50% | -5% |
| Dengue Fever | ~20% | ~12% | -8% |
| Hepatitis | ~30% | ~28% | -2% |

**Findings:**
- **Shingles** có intention-actual gap lớn nhất — nhiều khách hàng muốn tiêm nhưng chưa thực hiện trong lần đến phòng khám, chủ yếu do lo ngại chi phí và thiếu thông tin về lịch tiêm
- **Lý do đến tiêm vaccine** phổ biến nhất: (1) Khuyến nghị từ bác sĩ, (2) Tự tìm hiểu thông tin, (3) Người thân/bạn bè gợi ý
- Tháng 3 (M3) ghi nhận tỷ lệ chuyển đổi cao hơn M1 — cho thấy hiệu quả tích lũy của các hoạt động truyền thông

---

### 4.4 Shingles Intention — Phân tích chuyên sâu Vaccine Zona

**Key KPIs của Vaccine Shingles:**

| Chỉ số | Mô tả |
|--------|-------|
| % Khách hàng có ý định tiêm Shingles | Tỷ lệ trong tổng số khách 50+ |
| % Được HCP khuyến nghị | Tỷ lệ bác sĩ chủ động gợi ý Shingrix |
| % Tiêm mũi đầu tiên (Initial dose) | Tỷ lệ bắt đầu liệu trình lần đầu |
| Thời gian trung bình giữa 2 mũi | Trung bình số tháng giữa mũi 1 và mũi 2 |

**Findings:**
- Phần lớn khách tiêm Shingles đến trong tình huống **bác sĩ chủ động khởi xướng cuộc trò chuyện** — đây là điểm khác biệt quan trọng so với các vaccine khác mà khách tự chủ động hỏi
- Nhóm khách hàng có **bệnh lý nền** (tiểu đường, tim mạch, suy giảm miễn dịch) có tỷ lệ được tư vấn Shingles cao hơn đáng kể
- Thời gian giữa mũi 1 và mũi 2 thực tế **dài hơn so với khuyến cáo** — cần cải thiện hệ thống nhắc lịch tái tiêm

---

### 4.5 Clinic Introduction — Hiệu quả Khuyến nghị tại điểm bán

**Chỉ số đo lường hiệu quả HCP Recommendation:**

| Chỉ số | Ý nghĩa |
|--------|---------|
| % No Intention but Recommended | Tỷ lệ khách không có ý định nhưng được giới thiệu |
| % No Intention but Actual Get | Tỷ lệ chuyển đổi thành công từ không ý định → tiêm |
| % Rejection Rate | Tỷ lệ từ chối khuyến nghị và lý do |

**Lý do từ chối khuyến nghị phổ biến:**
- Chi phí vaccine quá cao
- Muốn tham khảo thêm ý kiến từ bác sĩ khác
- Không có triệu chứng nên chưa thấy cần thiết
- Sợ tác dụng phụ

**Người khuyến nghị hiệu quả nhất tại phòng khám:** Bác sĩ tư vấn trực tiếp > Dược sĩ > Điều dưỡng

---

### 4.6 Clinic Communication — Nhận biết & Truyền thông

**Findings:**
- Vật liệu truyền thông tại điểm bán (poster, brochure) có mức độ nhận diện thương hiệu **trung bình** — cần cải thiện về vị trí đặt và nội dung thông điệp
- Kênh **VNVC** có mức độ nhận biết thương hiệu vaccine cao hơn **Long Châu** do đầu tư mạnh hơn vào truyền thông phòng khám
- Tần suất tiếp xúc với thông tin vaccine online (mạng xã hội, website) chưa được tối ưu để chuyển đổi sang hành vi tại điểm tiêm

---

## 5. Recommendations

### 💊 Chiến lược Vaccine & Sản phẩm
- **Ưu tiên tập trung vào Shingles và Pneumococcal** — hai nhóm có intention-actual gap lớn nhất và tiềm năng tăng trưởng cao nhất trong nhóm 50+
- Phát triển **gói tiêm chủng combo** cho nhóm 50+ (Shingles + Pneumococcal + Influenza) để tăng giá trị đơn hàng và giảm rào cản về chi phí từng mũi lẻ
- Tăng cường **chuỗi cung ứng vaccine Shingrix và Prevenar** tại điểm bán — tình trạng hết hàng là nguyên nhân số 1 dẫn đến thay đổi thương hiệu không mong muốn

### 🏥 Tối ưu hóa vai trò HCP (Healthcare Professional)
- Đào tạo và hỗ trợ bác sĩ **chủ động khởi xướng cuộc trò chuyện về vaccine** với bệnh nhân 50+ — đặc biệt là nhóm có bệnh lý nền
- Xây dựng **kịch bản tư vấn chuẩn (HCP script)** cho từng nhóm vaccine, tập trung vào việc giải quyết lo ngại về chi phí và tác dụng phụ
- Triển khai **hệ thống nhắc lịch tái tiêm** (reminder system) cho vaccine Shingles để giảm khoảng cách giữa mũi 1 và mũi 2

### 📣 Truyền thông & Marketing tại phòng khám
- Cải thiện **chất lượng và vị trí vật liệu POS** (Point-of-Sale): đặt tại phòng chờ và phòng tư vấn thay vì chỉ tại quầy lễ tân
- Thiết kế nội dung truyền thông **ngắn gọn, dễ hiểu, tập trung vào lợi ích sức khỏe** thay vì thông tin kỹ thuật về vaccine
- Phối hợp chiến dịch **online-to-offline**: dùng mạng xã hội và Google Ads để tạo awareness, chuyển đổi thực tế tại phòng khám qua HCP

### 📅 Kế hoạch theo dõi & Báo cáo
- Duy trì khảo sát exit interview **hàng tháng** để theo dõi xu hướng intention-actual gap theo thời gian thực
- Thiết lập **dashboard tự động cập nhật hàng ngày** trên Power BI Service để phòng marketing theo dõi KPI kịp thời
- Mở rộng phạm vi khảo sát sang **thêm kênh phân phối** (bệnh viện tư, phòng khám đa khoa) để có bức tranh toàn diện hơn về thị trường vaccine

---

## 📊 Power BI Dashboard

Dashboard bao gồm **6 trang báo cáo** tương tác, với bộ lọc xuyên suốt theo Tháng, Kênh phân phối, Khu vực, Nhóm tuổi, và Nhóm bệnh:

| Trang | Tiêu đề | Chỉ tiêu chính |
|-------|---------|----------------|
| 1 | **Vaccine Units** | Tổng mũi tiêm, cơ cấu nhóm bệnh, xu hướng theo tháng |
| 2 | **Vaccination Share** | Thị phần hãng dược, donut chart theo nhóm bệnh |
| 3 | **Intention to Clinic** | Intention vs. Actual gap, lý do đến tiêm |
| 4 | **Shingles Intention** | KPIs Shingles, vai trò HCP, bệnh lý nền |
| 5 | **Clinic Introduction** | Tỷ lệ chuyển đổi từ khuyến nghị, lý do từ chối |
| 6 | **Clinic Communication** | Nhận biết thương hiệu, hiệu quả vật liệu truyền thông |

---

## 📄 License

Dự án này phục vụ mục đích portfolio và nghiên cứu. Dữ liệu được thu thập từ khảo sát exit interview tại phòng khám VNVC và Long Châu (M1–M3/2025). Mọi thông tin cá nhân đã được ẩn danh hóa trước khi xử lý.
