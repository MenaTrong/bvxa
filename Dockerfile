# Sử dụng hình ảnh OpenJDK 14
FROM openjdk:17-jdk-alpine

# Sao chép tất cả các tệp và thư mục cần thiết từ máy cục bộ vào container
COPY . /usr/src/demo

# Thiết lập thư mục làm việc
WORKDIR /usr/src/demo

EXPOSE 8080

# Chạy ứng dụng khi container được khởi động
CMD ["java", "-jar", "demo.jar"]
