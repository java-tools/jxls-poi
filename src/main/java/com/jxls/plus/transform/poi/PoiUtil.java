package com.jxls.plus.transform.poi;

import org.apache.poi.ss.usermodel.*;

/**
 * POI utility methods
 * @author Leonid Vysochyn
 */
public class PoiUtil {
    public static void setCellComment(Cell cell, String commentText, String commentAuthor, ClientAnchor anchor){
        Sheet sheet = cell.getSheet();
        Workbook wb = sheet.getWorkbook();
        Drawing drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        if( anchor == null ){
            anchor = factory.createClientAnchor();
            anchor.setCol1(0);
            anchor.setCol2(1);
            anchor.setRow1(0);
            anchor.setRow2(1);
        }
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(commentAuthor != null ? commentAuthor : "");
        cell.setCellComment( comment );
    }

    public WritableCellValue hyperlink(String address, String link, String linkTypeString){
        return new WritableHyperlink(address, link, linkTypeString);
    }

    public WritableCellValue hyperlink(String address, String linkTypeString){
        return new WritableHyperlink(address, linkTypeString);
    }

    public static void copySheetProperties(Sheet src, Sheet dest){
        dest.setAutobreaks(src.getAutobreaks());
        dest.setDisplayGridlines(src.getDisplayGuts());
        dest.setVerticallyCenter(src.getVerticallyCenter());
        dest.setFitToPage(src.getFitToPage());
        dest.setForceFormulaRecalculation(src.getForceFormulaRecalculation());
        dest.setRowSumsRight(src.getRowSumsRight());
        dest.setRowSumsBelow( src.getRowSumsBelow() );
        copyPrintSetup(src, dest);
    }

    private static void copyPrintSetup(Sheet src, Sheet dest) {
        PrintSetup srcPrintSetup = src.getPrintSetup();
        PrintSetup destPrintSetup = dest.getPrintSetup();
        destPrintSetup.setCopies(srcPrintSetup.getCopies());
        destPrintSetup.setDraft(srcPrintSetup.getDraft());
        destPrintSetup.setFitHeight(srcPrintSetup.getFitHeight());
        destPrintSetup.setFitWidth(srcPrintSetup.getFitWidth());
        destPrintSetup.setFooterMargin(srcPrintSetup.getFooterMargin());
        destPrintSetup.setHeaderMargin(srcPrintSetup.getHeaderMargin());
        destPrintSetup.setHResolution(srcPrintSetup.getHResolution());
        destPrintSetup.setLandscape(srcPrintSetup.getLandscape());
        destPrintSetup.setLeftToRight(srcPrintSetup.getLeftToRight());
        destPrintSetup.setNoColor(srcPrintSetup.getNoColor());
        destPrintSetup.setNoOrientation(srcPrintSetup.getNoOrientation());
        destPrintSetup.setNotes(srcPrintSetup.getNotes());
        destPrintSetup.setPageStart(srcPrintSetup.getPageStart());
        destPrintSetup.setPaperSize(srcPrintSetup.getPaperSize());
        destPrintSetup.setScale(srcPrintSetup.getScale());
        destPrintSetup.setUsePage(srcPrintSetup.getUsePage());
        destPrintSetup.setValidSettings(srcPrintSetup.getValidSettings());
        destPrintSetup.setVResolution( srcPrintSetup.getVResolution() );
    }
}
