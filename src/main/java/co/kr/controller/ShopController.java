package co.kr.controller;

import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.client.RestTemplate;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * Handles requests for the application home page.
 */
@Controller
public class ShopController {

	private static final Logger logger = LoggerFactory.getLogger(ShopController.class);
	List<Object> jsonlist =  new ArrayList<Object>();
	/**
	 * Simply selects the home view to render by returning its name.
	 */
	@RequestMapping(value = "orders" , method = RequestMethod.GET)
	public String shop(Locale locale, Model model, HttpServletRequest request) {
		try {
			logger.info("shop call {}.", locale);
			String authorityId = request.getParameter("authorityid");

			RestTemplate restTemplate = new RestTemplate();
			ResponseEntity<String> responseEntity = restTemplate.getForEntity("http://127.0.0.1:9220/controller/api/orders?authorityid="+authorityId, String.class);

			ObjectMapper mapper = new ObjectMapper();

			jsonlist = mapper.readValue(responseEntity.getBody(), new TypeReference<ArrayList<Object>>(){});

			model.addAttribute("totalData", jsonlist);
		} catch (Exception e) {
			// TODO: handle exception
		}

		return "orders";
	}

	@RequestMapping(value="/downLoadExcel.do")
	public void downLoadExcel(HttpServletResponse response, @RequestParam("fileName") String fileName) throws Exception {
		HSSFWorkbook objWorkBook = new HSSFWorkbook();
		HSSFSheet objSheet = null;
		HSSFRow objRow = null;
		HSSFCell objCell = null;       //셀 생성

		//제목 폰트
		HSSFFont font = objWorkBook.createFont();
		font.setFontHeightInPoints((short)9);
		font.setBoldweight((short)font.BOLDWEIGHT_BOLD);
		font.setFontName("맑은고딕");

		//제목 스타일에 폰트 적용, 정렬
		HSSFCellStyle styleHd = objWorkBook.createCellStyle();    //제목 스타일
		styleHd.setFont(font);
		styleHd.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		styleHd.setVerticalAlignment (HSSFCellStyle.VERTICAL_CENTER);
		objSheet = objWorkBook.createSheet("주문 현황 시트");     //워크시트 생성

		// 1행 헤더 지정
		String[] hearName = {"주문 식별번호","고객 식별번호","고객 이름","상품 식별번호","상품명"};
		objRow = objSheet.createRow(0);
		objRow.setHeight ((short) 0x150);

		for (int i = 0; i < hearName.length; i++) {
			objCell = objRow.createCell(i);
			objCell.setCellValue(hearName[i]);
			objCell.setCellStyle(styleHd);
		}
		
		// 2행 값 입력
		for (int i = 0; i < jsonlist.size(); i++) {
			objRow = objSheet.createRow(i+1);
			objRow.setHeight ((short) 0x150);

			List<String> totalList = getSplit(jsonlist.get(i));

			for (int i2 = 0; i2 < totalList.size(); i2++) {
				objCell = objRow.createCell(i2);
				objCell.setCellValue(totalList.get(i2));
				objCell.setCellStyle(styleHd);
			}
		}


		response.setContentType("Application/Msexcel");
		response.setHeader("Content-Disposition", "ATTachment; Filename="+URLEncoder.encode(fileName,"UTF-8")+".xls");

		OutputStream fileOut  = response.getOutputStream();
		objWorkBook.write(fileOut);
		fileOut.close();

		response.getOutputStream().flush();
		response.getOutputStream().close();
	}

	public List<String> getSplit(Object object) {
		List<String> data = new ArrayList<String>();
		
		String result = object.toString().trim();
		result = result.replaceAll("[{}]", "");
		String[] infos = result.split(",");

		for (int i = 0; i < infos.length; i++) {
			String str = "";
			for (int j = infos[i].length() -1; j >= 0; j--) {
				if(infos[i].charAt(j) != '=') {
					str += infos[i].charAt(j);
				}
				else {
					break;
				}
			}
			
			str = new StringBuffer(str).reverse().toString();
			data.add(str);
		}
		return data;
	}


}
