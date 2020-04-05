package com.dealExcel.ggg;

import com.dealExcel.service.dealExcelservice;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.util.ArrayList;

/**
 * Created by IntelliJ IDEA.
 *
 * @Auther: jkMa
 * @Date: 2020/4/4 20:11
 * @ItemName: dealExcel
 */
@Controller
public class IndexController {

    @Autowired
    dealExcelservice dealEx;

    @RequestMapping("/index")
    public String toIndex() {
        System.out.println("跳转到导入文件的html页面");
        return "notfshed";
    }

    @ResponseBody
    @RequestMapping("/get_finished_excel")
    public void getAlreadyfinish(@RequestParam("finishedExc") MultipartFile finishedExc, @RequestParam("allpeople") MultipartFile allpeople, HttpServletResponse response, HttpServletRequest request) throws Exception {
        //   String webPath=request.getServletContext().getRealPath("");
        String filePath = "C:\\Dfile\\myfile\\";
        File file1 = new File(filePath);
        file1.mkdirs();
        finishedExc.transferTo(new File(filePath + finishedExc.getOriginalFilename()));
        File finishfilepath = new File(filePath + finishedExc.getOriginalFilename());
        ArrayList<String> fishedList = dealEx.getDetil(finishfilepath);

        allpeople.transferTo(new File(filePath + allpeople.getOriginalFilename()));
        File allpeoplefilepath = new File(filePath + allpeople.getOriginalFilename());
        ArrayList<String> allpeoList = dealEx.getAllpeople(allpeoplefilepath);

        ArrayList<String> notfshedList = dealEx.notFinished(fishedList, allpeoList);
        System.out.println( "-未完成的同学====="+notfshedList.toString());

        dealEx.saveNotFished(notfshedList,response);
        // System.out.println("文件路径"+file.getOriginalFilename()+"---"+file.getInputStream());
    }
}
