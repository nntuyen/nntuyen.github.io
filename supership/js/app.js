document.addEventListener('DOMContentLoaded', (event) => {
  let data = [];
  for (let i = 0; i < 100; i++) data.push([]);

  let srcProvinces = dbProvinces.map(item => item["nameCode"]);
  let srcDistricts = dbDistricts.map(item => item["nameCode"]);
  let srcWards = dbWards.map(item => item["nameCode"]);

  jspreadsheet(document.getElementById('spreadsheet'), {
    data: data,
    onchange: onCellChanged,
    columns: [
      { type: 'text', title: 'STT', width: 30 },
      { type: 'text', title: 'Mã Đơn Khách Hàng', width: 120 },
      { type: 'text', title: 'Tên Người Nhận', width: 120 },
      { type: 'text', title: 'SĐT Người Nhận', width: 120 },
      { type: 'text', title: 'Địa chỉ', width: 500 },
      { type: 'dropdown', title: 'Tỉnh Thành Phố', width: 200, source: srcProvinces },
      { type: 'dropdown', title: 'Quận Huyện', width: 200, source: srcDistricts },
      { type: 'dropdown', title: 'Phường Xã', width: 200, source: srcWards },
      { type: 'text', title: 'Sản Phẩm', width: 120 },
      { type: 'text', title: 'Khối Lượng', width: 120 },
      { type: 'text', title: 'Thu Hộ', width: 120 },
      { type: 'text', title: 'Ghi Chú', width: 120 },
      { type: 'text', title: 'Người Trả Phí', width: 120 },
      { type: 'text', title: 'Gói Dịch Vụ', width: 120 },
      { type: 'text', title: 'Xem Thử Hàng', width: 120 },
      { type: 'text', title: 'Trị Giá', width: 120 },
      { type: 'text', title: 'Mã Khuyến Mãi', width: 120 },
      { type: 'text', title: 'Đổi Lấy Hàng Về', width: 120 }
    ]
  });
});

const onCellChanged = (instance, cell, x, y, value) => {
  if (x != 4) return;
  let row = Number(y) + 1;
  if (value == "") {
    $(instance).jexcel("setValue", "F" + row, "");
    $(instance).jexcel("setValue", "G" + row, "");
    $(instance).jexcel("setValue", "H" + row, "");
    return;
  }

  let result = parseAddress(value);
  $(instance).jexcel("setValue", "F" + row, result.province["nameCode"]);
  $(instance).jexcel("setValue", "G" + row, result.district["nameCode"]);
  $(instance).jexcel("setValue", "H" + row, result.ward["nameCode"]);
};

const toLatinText = str => {
  str = str.toLowerCase();
  str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  str = str.replace(/đ/g, "d");
  str = str.replace(/\u0300|\u0301|\u0303|\u0309|\u0323/g, "");
  str = str.replace(/\u02C6|\u0306|\u031B/g, "");
  return str;
}

const parseAddress = inputAddress => {
  let address = toLatinText(inputAddress);
  address = address.replace("hcm", "ho chi minh");

  let foundedProvince = "";
  let foundedDistrict = "";
  let foundedWard = "";
  for (let province of dbProvinces) {
    let provinceName = toLatinText(province["name"]);
    if (provinceName.startsWith("tinh ")) {
      provinceName = provinceName.replace("tinh_", "");
    } else if (provinceName.startsWith("thanh pho ")) {
      provinceName = provinceName.replace("thanh pho ", "");
    }

    if (address.includes(provinceName)) {
      foundedProvince = province;
      address = address.replace(provinceName, "");
      break;
    }
  }

  if (foundedProvince != "") {
    let filterDistricts = dbDistricts.filter(item => item["provinceId"] == foundedProvince["id"]);
    for (let district of filterDistricts) {
      let districtName = toLatinText(district["name"]);
      if (districtName.startsWith("huyen ")) {
        districtName = districtName.replace("huyen ", "");
      } else if (districtName.startsWith("thanh pho ")) {
        districtName = districtName.replace("thanh pho ", "");
      } else if (districtName.startsWith("thi xa ")) {
        districtName = districtName.replace("thi xa ", "");
      } else if (districtName.startsWith("quan ") && !/\d/.test(districtName)) {
        districtName = districtName.replace("quan ", "");
      }

      if (address.includes(districtName)) {
        foundedDistrict = district;
        address = address.replace(districtName, "");
        break;
      }
    }
  }

  if (foundedDistrict != "") {
    let filterWards = dbWards.filter(item => item["districtId"] == foundedDistrict["id"]);

    for (let ward of filterWards) {
      let wardName = toLatinText(ward["name"]);
      if (wardName.startsWith("xa ")) {
        wardName = wardName.replace("xa ", "");
      } else if (wardName.startsWith("thi tran ")) {
        wardName = wardName.replace("thi tran ", "");
      } else if (wardName.startsWith("phuong ") && !/\d/.test(wardName)) {
        wardName = wardName.replace("phuong ", "");
      }

      if (address.includes(wardName)) {
        foundedWard = ward;
        break;
      }
    }
  }

  return {
    province: foundedProvince,
    district: foundedDistrict,
    ward: foundedWard
  };
};

const htmlTableToExcel = (type) => {
  var data = document.getElementsByTagName('table')[0];
  var excelFile = XLSX.utils.table_to_book(data, { sheet: "sheet1" });
  XLSX.write(excelFile, { bookType: type, bookSST: true, type: 'base64' });
  XLSX.writeFile(excelFile, 'supership_order.' + type);
};