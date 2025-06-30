package com.hcmc.itc.cdsservice.service;

import com.hcmc.itc.cdsservice.core.dto.queries.BundlePycBaoGia;
import com.hcmc.itc.cdsservice.core.utils.CdsStringUtils;
import com.hcmc.itc.cdsservice.core.utils.NumToViet;
import com.hcmc.itc.cdsservice.core.utils.export.ExportFileDataBundlePojo;
import com.hcmc.itc.cdsservice.core.utils.words.ReplaceTextInWord;
import com.hcmc.itc.cdsservice.core.utils.words.XWPFDocumentUtils;
import com.hcmc.itc.cdsservice.core.utils.words.replacer.TextReplacer;
import com.hcmc.itc.cdsservice.core.xwpf.WordTagTemplate;
import com.hcmc.itc.cdsservice.feign.fallback.m1.M1PhieuYeuCauFeignServiceDelegate;
import com.hcmc.itc.cdsservice.feign.fallback.m31.M31EtcFeignServiceDelegate;
import com.hcmc.itc.cdsservice.feign.fallback.m31.M31KhangFeignServiceDelegate;
import com.hcmc.itc.cdsservice.feign.fallback.m33.M33DeviceCommonFeignServiceDelegate;
import com.hcmc.itc.cdsservice.feign.fallback.m33.M33ExperimentDeviceFeignServiceDelegate;
import com.hcmc.itc.cdsservice.feign.fallback.m33.M33QuyetToanFeignServiceDelegate;
import com.hcmc.itc.core.dto.response.queries.m1pyc.PhieuYeuCauDto;
import com.hcmc.itc.core.dto.response.queries.m31etc.NhanVienDto;
import com.hcmc.itc.core.dto.response.queries.m31khang.KhachHangDto;
import com.hcmc.itc.core.dto.response.queries.m33quyettoan.DmCongViecDto;
import com.hcmc.itc.core.dto.response.queries.m33quyettoan.DonGiaNgoaiNganhDto;
import com.hcmc.itc.core.dto.response.queries.m33quyettoan.DuToanDto;
import com.hcmc.itc.core.utils.CdsTndCoreUtils;
import com.hcmc.itc.core.utils.ETypeBaoGia;
import com.hcmc.itc.core.utils.ETypeExport;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.*;

/**
 * @author {thuanpk, quyenhvt, thuannp, hoan2nt, dan2nqt}@hcmpc.com.vn
 */
@Slf4j
@Component
@RequiredArgsConstructor
public class ProcessBaoGia {

    @Value("${etc.template.dirPath}")
    private String pathTemplateWordDirPath;
    @Value("${etc.template.bangBaoGiaPyc.filename}")
    private String pathTemplateWordFilename;
    @Value("${etc.export.pathForQRCode}")
    private String pathForQRCode;
    @Value("${etc.export.gotenberg.serviceHostGotenberg}")
    private String serviceHostGotenberg;
    @Value("${etc.export.gotenberg.convertToPdfUri}")
    private String convertToPdfUri;

    private static final int FONT_SIZE = 11;

    private final M1PhieuYeuCauFeignServiceDelegate m1PhieuYeuCauFeignServiceDelegate;
    private final M31KhangFeignServiceDelegate m31KhangFeignServiceDelegate;
    private final M33ExperimentDeviceFeignServiceDelegate m33ExperimentDeviceFeignServiceDelegate;
    private final M33DeviceCommonFeignServiceDelegate m33DeviceCommonFeignServiceDelegate;
    private final M33QuyetToanFeignServiceDelegate m33QuyetToanFeignServiceDelegate;
    private final M31EtcFeignServiceDelegate m31EtcFeignServiceDelegate;
    private final ReplaceTextInWord replaceTextInWord;
    private final TextReplacer textReplacer;

    public ByteArrayInputStream processExportBaoGiaPyc(ExportFileDataBundlePojo bundlePojo, ETypeExport typeExport) {
        // Get M1.phieu_yeu_cau by id
        BundlePycBaoGia bundlePycBaoGia = (BundlePycBaoGia)bundlePojo.getBundle();
        Long idPyc = bundlePycBaoGia.getIdPyc();
        ETypeBaoGia eTypeBaoGia = ETypeBaoGia.retrieveByCode(bundlePycBaoGia.getTypeBaoGia());

        PhieuYeuCauDto phieuYeuCauDto = m1PhieuYeuCauFeignServiceDelegate.getPhieuYeuCauById(idPyc);
        log.info("processExportBaoGiaPyc -> phieuYeuCauDto: {}", phieuYeuCauDto);
        if (phieuYeuCauDto == null) {
            phieuYeuCauDto = m1PhieuYeuCauFeignServiceDelegate.getPhieuYeuCauById(idPyc);
            if (phieuYeuCauDto == null) {
                log.warn("[ProcessBaoGia] - processExportBaoGiaPyc -> phieuYeuCauDto is null -> bundlePojo: {}, typeExport: {}", bundlePojo, typeExport);
                return new ByteArrayInputStream(new byte[0]);
            }
        }

        // Get M3.1.khach_hang by maKhang
        KhachHangDto khachHangDto = m31KhangFeignServiceDelegate.feignGetKhangById(phieuYeuCauDto.getIdKhang());
        log.info("processExportBaoGiaPyc -> khachHangDto: {}", khachHangDto);
        if (khachHangDto == null) {
            khachHangDto = m31KhangFeignServiceDelegate.feignGetKhangById(phieuYeuCauDto.getIdKhang());
            if (khachHangDto == null) {
                log.warn("[ProcessBaoGia] - processExportBaoGiaPyc -> khachHangDto is null -> bundlePojo: {}, typeExport: {}", bundlePojo, typeExport);
                return new ByteArrayInputStream(new byte[0]);
            }
        }

        // Get data du toan
        List<DuToanDto> duToanDtos = m33QuyetToanFeignServiceDelegate.feignGetDutoansByIdPyc(idPyc);
        log.info("processExportBaoGiaPyc -> duToanDtos: {}", duToanDtos);
        if (duToanDtos == null || duToanDtos.size() == 0) {
            duToanDtos = m33QuyetToanFeignServiceDelegate.feignGetDutoansByIdPyc(idPyc);
            if (duToanDtos == null || duToanDtos.size() == 0) {
                log.warn("[ProcessBaoGia] - processExportBaoGiaPyc -> bangCtinhKhangNgoaiDtos NOT FOUND -> phieuYeuCauDto: {}", phieuYeuCauDto);
                return new ByteArrayInputStream(new byte[0]);
            }
        }

        String pathFile = pathTemplateWordDirPath + "/" + pathTemplateWordFilename;
        //String pathFile = "E:\\sourceCode\\repository\\hcmc-microservice\\cds-tnd-docx-template\\cds-tnd-mau-bao-gia-pyc.docx";
        try (
                XWPFDocument document = new XWPFDocument(new FileInputStream(pathFile));
                ByteArrayOutputStream out = new ByteArrayOutputStream();
        ) {
            duToanDtos.sort(Comparator.comparingLong(DuToanDto::getIdDgiaNgoaiNganh));
            // Pargagraph
            Map<String, String> fieldsBody = new HashMap<>();
            // 1. Info
            fieldsBody.put("ten_ctrinh", CdsTndCoreUtils.blankWithDefaultValue(phieuYeuCauDto.getTenCongTrinh(), "-"));
            fieldsBody.put("ten_khang", CdsTndCoreUtils.blankWithDefaultValue(khachHangDto.getTenKhang(), "-"));
            // 2. Nhân viên lập báo giá
            NhanVienDto nvienLapBaoGia = m31EtcFeignServiceDelegate.getNhanVienById(phieuYeuCauDto.getIdNvienXnYc());
            if (nvienLapBaoGia == null) {
                nvienLapBaoGia = m31EtcFeignServiceDelegate.getNhanVienById(phieuYeuCauDto.getIdNvienKySoBaoGia());
            }
            if (nvienLapBaoGia != null) {
                fieldsBody.put("hten_nguoi_lap_bgia", CdsTndCoreUtils.blankWithDefaultValue(nvienLapBaoGia.getHoTen(), "-"));
                fieldsBody.put("sdt_nguoi_lap_bgia", CdsTndCoreUtils.blankWithDefaultValue(nvienLapBaoGia.getDienThoai(), "-"));
            }

            // Process table bao gia
            XWPFTable table = document.getTableArray(1);
            int rowIdx = 3;
            int indexRowCloneCphiTnghiem = 1;
            int indexRowCloneCviec = 2;
            int sttCviec = 1;
            DecimalFormat formatter = new DecimalFormat("###,###,###");
            DecimalFormat formatterHeSo = new DecimalFormat("###.###");
            BigDecimal tongCong = BigDecimal.ZERO;
            BigDecimal tongCphiTnghiem = BigDecimal.ZERO;
            BigDecimal tongCphiKhac = BigDecimal.ZERO;
            BigDecimal tienChietGiam = BigDecimal.ZERO;
            BigDecimal tongCphiTruocThue = BigDecimal.ZERO;
            BigDecimal tienThueGtgt = BigDecimal.ZERO;
            BigDecimal tongCphiSauThue = BigDecimal.ZERO;
            List<List<String>> cphiKhacs = new ArrayList<>();
            int sttChiPhiKhac = 0;
            for (DuToanDto duToanDto : duToanDtos) {
                List<String> childCphiKhacCells = new ArrayList<>();
                // Lay don-gia-ngoai-nganh theo idDonGiaNgoaiNganh
                DonGiaNgoaiNganhDto donGiaNgoaiNganhDto = m33QuyetToanFeignServiceDelegate.feignGetDonGiaNgoaiNganhById(duToanDto.getIdDgiaNgoaiNganh());
                if (donGiaNgoaiNganhDto == null) {
                    donGiaNgoaiNganhDto = m33QuyetToanFeignServiceDelegate.feignGetDonGiaNgoaiNganhById(duToanDto.getIdDgiaNgoaiNganh());
                }
                log.info("[ProcessBaoGia] - duToanDto: {}", duToanDto);
                log.info("[ProcessBaoGia] - donGiaNgoaiNganhDto: {}", donGiaNgoaiNganhDto);

                BigDecimal dgiaVchuyen = duToanDto.getDgiaVanChuyen() == null ? BigDecimal.ZERO : duToanDto.getDgiaVanChuyen();
                BigDecimal quangDuong = duToanDto.getQuangDuong() == null ? BigDecimal.ZERO : duToanDto.getQuangDuong();
                if (quangDuong.longValue() > 0) {
                    childCphiKhacCells.add((sttChiPhiKhac + 1) + "");
                    sttChiPhiKhac++;
                    childCphiKhacCells.add("");
                    childCphiKhacCells.add("Km");
                    childCphiKhacCells.add(duToanDto.getQuangDuong().toString());
                    childCphiKhacCells.add(formatter.format(duToanDto.getDgiaVanChuyen().longValue()).replace(",", "."));
                    childCphiKhacCells.add(formatter.format((quangDuong.multiply(dgiaVchuyen).setScale(0, RoundingMode.HALF_UP))).replace(",", "."));
                    cphiKhacs.add(childCphiKhacCells);
                }

                tongCphiKhac = tongCphiKhac.add(dgiaVchuyen.multiply(quangDuong)).setScale(0, RoundingMode.HALF_UP);
                if (donGiaNgoaiNganhDto != null) {
                    List<String> cviecCells = new ArrayList<>();
                    cviecCells.add(sttCviec + "");
                    cviecCells.add(donGiaNgoaiNganhDto.getMoTa());
                    cviecCells.add(donGiaNgoaiNganhDto.getTenDviTinh());
                    cviecCells.add(duToanDto.getSluongTnghiem().toString());
                    cviecCells.add(formatter.format(duToanDto.getDonGia().longValue()).replace(",", "."));
                    cviecCells.add(formatter.format((duToanDto.getTienTruocThue().setScale(0, RoundingMode.HALF_UP)).longValue()).replace(",", "."));
                    tongCphiTnghiem = tongCphiTnghiem.add((duToanDto.getTienTruocThue().setScale(0, RoundingMode.HALF_UP))
                            .subtract((dgiaVchuyen.multiply(quangDuong).setScale(0, RoundingMode.HALF_UP))));
                    XWPFDocumentUtils.cloneRow(table, indexRowCloneCviec, rowIdx++, cviecCells, false, FONT_SIZE);
                    sttCviec++;
                }
            }
            // Chi phí thử nghiệm
            List<String> cphiTnghiemCells = new ArrayList<>();
            cphiTnghiemCells.add("I");
            cphiTnghiemCells.add("Chi phí thử nghiệm");
            cphiTnghiemCells.add("");
            cphiTnghiemCells.add("");
            cphiTnghiemCells.add("");
            cphiTnghiemCells.add(formatter.format(tongCphiTnghiem.longValue()).replace(",", "."));
            XWPFDocumentUtils.cloneRow(table, indexRowCloneCphiTnghiem, 3, cphiTnghiemCells, false, FONT_SIZE);
            rowIdx++;

            // Thàn phần con của chi phí khác
            int indexCphiKhac = rowIdx;
            log.info("indexCphiKhac: {}", indexCphiKhac);

            for (List<String> childCphiKhacCells : cphiKhacs) {
                XWPFDocumentUtils.cloneRow(table, indexRowCloneCviec, rowIdx++, childCphiKhacCells, false, FONT_SIZE);
            }
            // Chi phí khác
            List<String> cphiKhacCells = new ArrayList<>();
            cphiKhacCells.add("II");
            cphiKhacCells.add("Chi phí khác");
            cphiKhacCells.add("");
            cphiKhacCells.add("");
            cphiKhacCells.add("");
            cphiKhacCells.add(formatter.format(tongCphiKhac.longValue()).replace(",", "."));
            XWPFDocumentUtils.cloneRow(table, indexRowCloneCphiTnghiem, indexCphiKhac, cphiKhacCells, false, FONT_SIZE);

            // Table
            Map<String, String> fieldsTable = new HashMap<>();
            fieldsTable.put("so_bban_bgia", CdsTndCoreUtils.blankWithDefaultValue(phieuYeuCauDto.getSoHieuPhieu(), "-").split("/")[0] + "/" + "BG-ETC");
            fieldsTable.put("ddiem_tao", "Tân Bình");
            fieldsTable.put("tgian_tao", CdsStringUtils.dateNgayThangNam(new Date()));
            // Tổng cộng
            tongCong = tongCphiTnghiem.add(tongCphiKhac);
            tienChietGiam = tongCong.multiply(phieuYeuCauDto.getTiLeChietGiam()).setScale(0, RoundingMode.HALF_UP);
            tongCphiTruocThue = tongCong.subtract(tienChietGiam);
            tienThueGtgt = tongCphiTruocThue.multiply(phieuYeuCauDto.getTiLeThueGtgt()).setScale(0, RoundingMode.HALF_UP);
            tongCphiSauThue = tongCphiTruocThue.add(tienThueGtgt).setScale(0, RoundingMode.HALF_UP);
            fieldsTable.put("tong_cong", formatter.format(tongCong.longValue()).replace(",", "."));
            fieldsTable.put("ti_le_cgiam", formatterHeSo.format(phieuYeuCauDto.getTiLeChietGiam().multiply(new BigDecimal("100"))).replace(".", ","));
            fieldsTable.put("gtien_cgiam", formatter.format((tienChietGiam.setScale(0, RoundingMode.HALF_UP)).longValue()).replace(",", "."));
            fieldsTable.put("gtien_cphi_tthue", formatter.format(tongCphiTruocThue.longValue()).replace(",", "."));
            fieldsTable.put("ti_le_thue_gtgt", formatterHeSo.format(phieuYeuCauDto.getTiLeThueGtgt().multiply(new BigDecimal("100"))).replace(".", ","));
            fieldsTable.put("gtien_thue_gtgt", formatter.format(tienThueGtgt.longValue()).replace(",", "."));
            fieldsTable.put("gtien_cphi_sthue", formatter.format(tongCphiSauThue.longValue()).replace(",", "."));
            fieldsTable.put("gtien_bang_chu", NumToViet.num2String(tongCphiSauThue.longValue()));

            // 1. Nhân viên ký số
            NhanVienDto nvienKySo = m31EtcFeignServiceDelegate.getNhanVienById(phieuYeuCauDto.getIdNvienKySoBaoGia());
            if (nvienKySo == null) {
                nvienKySo = m31EtcFeignServiceDelegate.getNhanVienById(phieuYeuCauDto.getIdNvienKySoBaoGia());
            }
            if (nvienKySo != null) {
                fieldsTable.put("ddien_etc_ky", CdsTndCoreUtils.blankWithDefaultValue(nvienKySo.getHoTen(), "-"));
            }

            if (phieuYeuCauDto.getTiLeChietGiam() == null || phieuYeuCauDto.getTiLeChietGiam().longValue() <= 0) {
                table.removeRow(rowIdx + 2);
            }

            WordTagTemplate.replaceParagraphs(document, fieldsBody);
            WordTagTemplate.replaceTables(document, fieldsTable);

//            String qrCodeContent = pathForQRCode;
//            KeyValuePair k = new KeyValuePair();
//            k.setKey("key");
//            k.setValue("WEB_HOST");
//            List<KeyValuePair> kv = new ArrayList<>();
//            kv.add(k);
//            List<TblParmetersDto> urlsQrCode = m33DeviceCommonFeignServiceDelegate.feignFilterTblParmeters(kv);
//            if (urlsQrCode != null && !urlsQrCode.isEmpty()) {
//                qrCodeContent = urlsQrCode.get(0).getValueData();
//                qrCodeContent += "/file/preview?id=" + idPyc + "&docType=baogia";
//            }
//            XWPFDocumentUtils.createQRCodeInDocument(document, qrCodeContent, 100, 100);

            table.removeRow(1);
            table.removeRow(1);

            document.write(out);
            document.close();

            switch (typeExport) {
                case DOCX:
                    return new ByteArrayInputStream(out.toByteArray());
                case PDF:
                    String fileName = CdsTndCoreUtils.blankWithDefaultValue(pathTemplateWordFilename, "bao_gia_pyc.docx");
                    String mimeType = CdsStringUtils.dectectMimeTypeFromFileName(fileName);
                    byte[] outBytePdf = CdsStringUtils.convertToPdfByGotenberg(serviceHostGotenberg, convertToPdfUri, fileName, out.toByteArray(), mimeType);
                    return new ByteArrayInputStream(outBytePdf);
            }
        } catch (Exception e) {
            log.error("export file bao-gia error", e);
        }

        return new ByteArrayInputStream(new byte[0]);
    }

    private List<String> genDataVttbForRowTable(int stt, String tenDmDtuongTnghiem) {

        List<String> vttbValues = new ArrayList<>();
        // Ten đối tượng thử nghiệm
        String tenDmDoiTuongTnghiem = CdsTndCoreUtils.blankWithDefaultValue(tenDmDtuongTnghiem, "-");
        vttbValues.add(stt + ". " + tenDmDoiTuongTnghiem);

        return vttbValues;
    }

    private List<String> genDataDmCongViecForRowTable(int sttDmCviec, DmCongViecDto cviec) {
        // Get row 1
        List<String> vttbValues = new ArrayList<>();
        BigDecimal tong = BigDecimal.ZERO;
        // Stt
        vttbValues.add(sttDmCviec + "");
        // Ma
        vttbValues.add(CdsTndCoreUtils.blankWithDefaultValue(cviec.getMa(), "-"));
        // Ten
        vttbValues.add(CdsTndCoreUtils.blankWithDefaultValue(cviec.getTen(), "-"));
        // Nhan-cong
        BigDecimal nhanCong = cviec.getCphiNcong() == null ? BigDecimal.ZERO : cviec.getCphiNcong();
        nhanCong = nhanCong.setScale(2, RoundingMode.HALF_UP);
        vttbValues.add(CdsTndCoreUtils.concurrencyToVietnamese(nhanCong));
        tong = tong.add(nhanCong);
        // Vat lieu
        BigDecimal vatLieu = cviec.getCphiVlieu() == null ? BigDecimal.ZERO : cviec.getCphiVlieu();
        vatLieu = vatLieu.setScale(2, RoundingMode.HALF_UP);
        vttbValues.add(CdsTndCoreUtils.concurrencyToVietnamese(vatLieu));
        tong = tong.add(vatLieu);
        // May thi cong
        BigDecimal mayThiCong = cviec.getCphiMayTcong() == null ? BigDecimal.ZERO : cviec.getCphiMayTcong();
        mayThiCong = mayThiCong.setScale(2, RoundingMode.HALF_UP);
        vttbValues.add(CdsTndCoreUtils.concurrencyToVietnamese(mayThiCong));
        tong = tong.add(mayThiCong);
        // Don-gia
        tong = tong.setScale(2, RoundingMode.HALF_UP);
        vttbValues.add(CdsTndCoreUtils.concurrencyToVietnamese(tong));

        return vttbValues;
    }
}
