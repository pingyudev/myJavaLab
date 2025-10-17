package com.example.myjavalab.utils;

/**
 * 书签范围类，用于表示书签在文档中的位置范围
 */
public class BookmarkRange {
    private final int start;
    private final int end;
    
    /**
     * 构造函数
     * @param start 书签起始位置（段落索引）
     * @param end 书签结束位置（段落索引）
     */
    public BookmarkRange(int start, int end) {
        this.start = start;
        this.end = end;
    }
    
    /**
     * 获取书签起始位置
     * @return 起始段落索引
     */
    public int getStart() {
        return start;
    }
    
    /**
     * 获取书签结束位置
     * @return 结束段落索引
     */
    public int getEnd() {
        return end;
    }
    
    /**
     * 检查书签是否有效（起始位置小于等于结束位置）
     * @return 如果书签范围有效返回true，否则返回false
     */
    public boolean isValid() {
        return start >= 0 && end >= 0 && start <= end;
    }
    
    /**
     * 检查书签是否未找到（起始位置为-1）
     * @return 如果书签未找到返回true，否则返回false
     */
    public boolean isNotFound() {
        return start == -1 && end == -1;
    }
    
    @Override
    public String toString() {
        if (isNotFound()) {
            return "BookmarkRange{NOT_FOUND}";
        }
        return "BookmarkRange{start=" + start + ", end=" + end + "}";
    }
    
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (obj == null || getClass() != obj.getClass()) return false;
        BookmarkRange that = (BookmarkRange) obj;
        return start == that.start && end == that.end;
    }
    
    @Override
    public int hashCode() {
        return java.util.Objects.hash(start, end);
    }
}
