// 快取相關常數
const CACHE_EXPIRATION = 21600; // 快取時效 6 小時
const CACHE_KEYS = {
    LIMIT_OF_SCHOOLS: 'limitOfSchools',
    DEPARTMENT_OPTIONS: 'departmentOptions',
    EXAM_DATA: 'examData',
    CHOICES_DATA: 'choicesData',
};

/**
 * @description 將大型資料分段儲存到快取
 * @param {string} key - 快取鍵值
 * @param {Object} data - 要存入的資料
 */
function setChunkedCacheData(key, data) {
    const cache = CacheService.getScriptCache();
    const str = JSON.stringify(data);

    // 將資料分段，每段約 90KB
    const chunkSize = 90000;
    const chunks = [];
    for (let i = 0; i < str.length; i += chunkSize) {
        chunks.push(str.slice(i, i + chunkSize));
    }

    // 儲存分段數量
    const cacheObj = {};
    cacheObj[`${key}_chunks`] = chunks.length;

    // 儲存每個分段
    chunks.forEach((chunk, i) => {
        cacheObj[`${key}_${i}`] = chunk;
    });

    cache.putAll(cacheObj, CACHE_EXPIRATION);
}

/**
 * @description 從快取中讀取分段資料並組合
 * @param {string} key - 快取鍵值
 * @returns {Object|null} 快取資料或 null
 */
function getChunkedCacheData(key) {
    const cache = CacheService.getScriptCache();
    const numChunks = Number(cache.get(`${key}_chunks`));

    if (!numChunks) {
        return null;
    }

    // 讀取所有分段
    const keys = Array.from({ length: numChunks }, (_, i) => `${key}_${i}`);
    const chunks = cache.getAll(keys);

    if (!chunks || Object.keys(chunks).length === 0) {
        return null;
    }

    // 組合所有分段
    const jsonStr = Array.from(
        { length: numChunks },
        (_, i) => chunks[`${key}_${i}`]
    ).join('');

    try {
        return JSON.parse(jsonStr);
    } catch (e) {
        Logger.log('快取資料解析錯誤：%s', e.message);
        return null;
    }
}

/**
 * @description 設定快取資料（自動判斷是否需要分段）
 * @param {string} key - 快取鍵值
 * @param {Object} data - 要存入的資料
 */
function setCacheData(key, data) {
    const str = JSON.stringify(data);
    if (str.length > 90000) {
        setChunkedCacheData(key, data);
    } else {
        const cache = CacheService.getScriptCache();
        cache.put(key, str, CACHE_EXPIRATION);
    }
}

/**
 * @description 取得快取資料（自動判斷是否為分段資料）
 * @param {string} key - 快取鍵值
 * @returns {Object|null} 快取資料或 null
 */
function getCacheData(key) {
    const cache = CacheService.getScriptCache();
    const chunksCount = cache.get(`${key}_chunks`);

    if (chunksCount) {
        return getChunkedCacheData(key);
    }

    const data = cache.get(key);
    return data ? JSON.parse(data) : null;
}
