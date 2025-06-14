// 快取相關常數和安全性設定
const CACHE_EXPIRATION = 21600; // 快取時效 6 小時
const MAX_CACHE_EXPIRATION = 86400; // 最大快取時效 24 小時
const CHUNK_SIZE = 90000; // 每段約 90KB
const MAX_CACHE_SIZE = 1000000; // 最大快取大小 1MB
const MAX_CHUNKS = 50; // 最大分段數

const CACHE_KEYS = {
  LIMIT_OF_SCHOOLS: "limitOfSchools",
  DEPARTMENT_OPTIONS: "departmentOptions",
  EXAM_DATA: "examData",
  CHOICES_DATA: "choicesData",
  USER_DATA_PREFIX: "userData_",
};

/**
 * @description 驗證快取鍵值的安全性
 * @param {string} key - 快取鍵值
 * @returns {boolean} 是否為安全的鍵值
 */
function isValidCacheKey(key) {
  return (
    typeof key === "string" &&
    key.length > 0 &&
    key.length <= 100 &&
    /^[a-zA-Z0-9_-]+$/.test(key)
  );
}

/**
 * @description 驗證快取資料大小
 * @param {any} data - 要快取的資料
 * @returns {boolean} 資料大小是否合理
 */
function isValidCacheSize(data) {
  try {
    const jsonStr = JSON.stringify(data);
    return jsonStr.length <= MAX_CACHE_SIZE;
  } catch (error) {
    Logger.log("驗證快取大小時發生錯誤：%s", error.message);
    return false;
  }
}

/**
 * @description 清理過期或損壞的快取
 * @param {string} keyPrefix - 要清理的快取鍵值前綴
 */
function cleanupCache(keyPrefix) {
  try {
    const cache = CacheService.getScriptCache();

    // 嘗試清理可能的分段快取
    for (let i = 0; i < MAX_CHUNKS; i++) {
      const chunkKey = `${keyPrefix}_${i}`;
      cache.remove(chunkKey);
    }

    cache.remove(`${keyPrefix}_chunks`);
    cache.remove(keyPrefix);

    Logger.log("已清理快取：%s", keyPrefix);
  } catch (error) {
    Logger.log("清理快取時發生錯誤：%s", error.message);
  }
}

/**
 * @description 將大型資料分段儲存到快取（安全版本）
 * @param {string} key - 快取鍵值
 * @param {Object} data - 要存入的資料
 * @param {number} [expiration] - 快取時效（秒）
 */
function setChunkedCacheData(key, data, expiration = CACHE_EXPIRATION) {
  try {
    // 驗證輸入
    if (!isValidCacheKey(key)) {
      throw new Error("(setChunkedCacheData)無效的快取鍵值");
    }

    if (!data) {
      Logger.log("(setChunkedCacheData)快取資料為空，鍵值：%s", key);
      return;
    }

    // 驗證到期時間
    const validExpiration = Math.min(
      Math.max(expiration, 60),
      MAX_CACHE_EXPIRATION
    );

    const cache = CacheService.getScriptCache();
    const jsonStr = JSON.stringify(data);

    // 檢查資料大小
    if (jsonStr.length > MAX_CACHE_SIZE) {
      Logger.log("資料過大，無法快取：%s (%d 字元)", key, jsonStr.length);
      return;
    }

    // 清理可能存在的舊快取
    cleanupCache(key);

    // 分段處理
    const chunks = [];
    for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
      chunks.push(jsonStr.slice(i, i + CHUNK_SIZE));
    }

    if (chunks.length > MAX_CHUNKS) {
      Logger.log("分段數量過多，無法快取：%s (%d 段)", key, chunks.length);
      return;
    }

    // 準備快取資料
    const cacheObj = {};
    cacheObj[`${key}_chunks`] = chunks.length;

    // 儲存每個分段
    chunks.forEach((chunk, i) => {
      cacheObj[`${key}_${i}`] = chunk;
    });

    // 批次寫入快取
    cache.putAll(cacheObj, validExpiration);
    Logger.log(
      "已分段快取資料：%s (%d 段，%d 字元)",
      key,
      chunks.length,
      jsonStr.length
    );
  } catch (error) {
    Logger.log("設定分段快取時發生錯誤：%s", error.message);
    cleanupCache(key); // 清理可能損壞的快取
  }
}

/**
 * @description 從快取中讀取分段資料並組合（安全版本）
 * @param {string} key - 快取鍵值
 * @returns {Object|null} 快取資料或 null
 */
function getChunkedCacheData(key) {
  try {
    // 驗證輸入
    if (!isValidCacheKey(key)) {
      Logger.log("(getChunkedCacheData)無效的快取鍵值：%s", key);
      return null;
    }

    const cache = CacheService.getScriptCache();
    const numChunks = Number(cache.get(`${key}_chunks`));

    if (!numChunks || numChunks <= 0 || numChunks > MAX_CHUNKS) {
      return null;
    }

    // 讀取所有分段
    const keys = Array.from({ length: numChunks }, (_, i) => `${key}_${i}`);
    const chunks = cache.getAll(keys);

    if (!chunks || Object.keys(chunks).length !== numChunks) {
      Logger.log("快取分段不完整，清理：%s", key);
      cleanupCache(key);
      return null;
    }

    // 組合所有分段
    const jsonStr = Array.from(
      { length: numChunks },
      (_, i) => chunks[`${key}_${i}`]
    ).join("");

    if (!jsonStr) {
      Logger.log("快取資料為空：%s", key);
      return null;
    }

    const data = JSON.parse(jsonStr);
    Logger.log("成功讀取分段快取：%s (%d 段)", key, numChunks);
    return data;
  } catch (error) {
    Logger.log("讀取分段快取時發生錯誤：%s", error.message);
    cleanupCache(key); // 清理損壞的快取
    return null;
  }
}

/**
 * @description 設定快取資料（自動判斷是否需要分段，安全版本）
 * @param {string} key - 快取鍵值
 * @param {Object} data - 要存入的資料
 * @param {number} [expiration] - 快取時效（秒）
 */
function setCacheData(key, data, expiration = CACHE_EXPIRATION) {
  try {
    // 驗證輸入
    if (!isValidCacheKey(key)) {
      Logger.log("(setCacheData)無效的快取鍵值：%s", key);
      return;
    }

    if (!data) {
      Logger.log("(setCacheData)快取資料為空：%s", key);
      return;
    }

    // 驗證資料大小
    if (!isValidCacheSize(data)) {
      Logger.log("(setCacheData)快取資料過大：%s", key);
      return;
    }

    // 驗證到期時間
    const validExpiration = Math.min(
      Math.max(expiration, 60),
      MAX_CACHE_EXPIRATION
    );

    const jsonStr = JSON.stringify(data);

    if (jsonStr.length > CHUNK_SIZE) {
      // 使用分段快取
      setChunkedCacheData(key, data, validExpiration);
    } else {
      // 直接快取
      const cache = CacheService.getScriptCache();
      cache.put(key, jsonStr, validExpiration);
      Logger.log("已快取資料：%s (%d 字元)", key, jsonStr.length);
    }
  } catch (error) {
    Logger.log("設定快取時發生錯誤：%s", error.message);
  }
}

/**
 * @description 取得快取資料（自動判斷是否為分段資料，安全版本）
 * @param {string} key - 快取鍵值
 * @returns {Object|null} 快取資料或 null
 */
function getCacheData(key) {
  try {
    // 驗證輸入
    if (!isValidCacheKey(key)) {
      Logger.log("(getCacheData)無效的快取鍵值：%s", key);
      return null;
    }

    const cache = CacheService.getScriptCache();

    // 檢查是否為分段快取
    const chunksCount = cache.get(`${key}_chunks`);
    if (chunksCount) {
      return getChunkedCacheData(key);
    }

    // 讀取一般快取
    const jsonStr = cache.get(key);
    if (!jsonStr) {
      return null;
    }

    const data = JSON.parse(jsonStr);
    Logger.log("成功讀取快取：%s", key);
    return data;
  } catch (error) {
    Logger.log("讀取快取時發生錯誤：%s", error.message);
    // 清理可能損壞的快取
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
    } catch (cleanupError) {
      Logger.log("清理損壞快取失敗：%s", cleanupError.message);
    }
    return null;
  }
}

/**
 * @description 清除指定的快取資料
 * @param {string} key - 快取鍵值
 */
function removeCacheData(key) {
  try {
    if (!isValidCacheKey(key)) {
      Logger.log("(removeCacheData)無效的快取鍵值：%s", key);
      return;
    }

    cleanupCache(key);
    Logger.log("已移除快取：%s", key);
  } catch (error) {
    Logger.log("移除快取時發生錯誤：%s", error.message);
  }
}

/**
 * @description 清除所有快取資料
 */
function clearAllCache() {
  try {
    // 清除已知的快取鍵值
    for (const cacheKey of Object.values(CACHE_KEYS)) {
      cleanupCache(cacheKey);
    }

    Logger.log("已清除所有快取資料");
  } catch (error) {
    Logger.log("清除所有快取時發生錯誤：%s", error.message);
  }
}
