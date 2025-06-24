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
  USER_INDEX: "userIndex", // 使用者索引的快取鍵值
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

    Logger.log("(cleanupCache)已清理快取：%s", keyPrefix);
  } catch (error) {
    Logger.log("(cleanupCache)清理快取時發生錯誤：%s", error.message);
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
      Logger.log(
        "(setChunkedCacheData)資料過大，無法快取：%s (%d 字元)",
        key,
        jsonStr.length
      );
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
      Logger.log(
        "(setChunkedCacheData)分段數量過多，無法快取：%s (%d 段)",
        key,
        chunks.length
      );
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
      "(setChunkedCacheData)已分段快取資料：%s (%d 段，%d 字元)",
      key,
      chunks.length,
      jsonStr.length
    );
  } catch (error) {
    Logger.log(
      "(setChunkedCacheData)設定分段快取時發生錯誤：%s",
      error.message
    );
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
      Logger.log("(getChunkedCacheData)快取分段不完整，清理：%s", key);
      cleanupCache(key);
      return null;
    }

    // 組合所有分段
    const jsonStr = Array.from(
      { length: numChunks },
      (_, i) => chunks[`${key}_${i}`]
    ).join("");

    if (!jsonStr) {
      Logger.log("(getChunkedCacheData)快取資料為空：%s", key);
      return null;
    }

    const data = JSON.parse(jsonStr);
    Logger.log(
      "(getChunkedCacheData)成功讀取分段快取：%s (%d 段)",
      key,
      numChunks
    );
    return data;
  } catch (error) {
    Logger.log(
      "(getChunkedCacheData)讀取分段快取時發生錯誤：%s",
      error.message
    );
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
    Logger.log("(getCacheData)成功讀取快取：%s", key);
    return data;
  } catch (error) {
    Logger.log("(getCacheData)讀取快取時發生錯誤：%s", error.message);
    // 清理可能損壞的快取
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
    } catch (cleanupError) {
      Logger.log("(getCacheData)清理損壞快取失敗：%s", cleanupError.message);
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
    Logger.log("(removeCacheData)已移除快取：%s", key);
  } catch (error) {
    Logger.log("(removeCacheData)移除快取時發生錯誤：%s", error.message);
  }
}

/**
 * @description 清除所有快取資料
 * @returns {boolean} 是否成功清除所有快取
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    const userCacheKeys = getAllUserCacheKeys();

    // 清除固定的快取鍵值（先嘗試針對性清除）
    for (const cacheKey of Object.values(CACHE_KEYS)) {
      cleanupCache(cacheKey);
    }

    for (const userKey of userCacheKeys) {
      cleanupCache(userKey);
    }
    Logger.log("(clearAllCache)已清除所有固定快取鍵值和使用者快取鍵值");

    // 使用暴力方式：直接重置整個快取空間
    // 這將清除所有快取資料，包含以 userData_ 為前綴的使用者快取
    try {
      cache.removeAll([]); // 傳入空陣列會清除所有快取
      Logger.log("(clearAllCache)已重置所有快取資料（包含使用者資料快取）");
    } catch (resetError) {
      Logger.log("(clearAllCache)重置快取時發生錯誤：%s", resetError.message);
      return false;
    }

    Logger.log("(clearAllCache)已成功清除所有快取資料");
    return true;
  } catch (error) {
    Logger.log("(clearAllCache)清除所有快取時發生錯誤：%s", error.message);
    return false;
  }
}

/**
 * @description 內部清除所有快取實作，供 main.js 呼叫以避免函式名稱衝突
 * @returns {boolean} 是否成功清除所有快取
 */
function clearAllCacheInternal() {
  return clearAllCache(); // 呼叫本檔案中的 clearAllCache
}

/**
 * @description 將使用者添加到快取索引中
 * @param {string} email - 使用者電子郵件
 * @returns {boolean} 是否成功添加到索引
 */
function addUserToIndex(email) {
  try {
    if (!email) {
      Logger.log("(addUserToIndex)無效的電子郵件");
      return false;
    }

    const cache = CacheService.getScriptCache();
    const safeEmail = getSafeKeyFromEmail(email);

    // 從快取獲取當前索引
    let userIndex = [];
    const indexStr = cache.get(CACHE_KEYS.USER_INDEX);

    if (indexStr) {
      try {
        userIndex = JSON.parse(indexStr);
      } catch (e) {
        Logger.log("(addUserToIndex)解析使用者索引發生錯誤：%s", e.message);
        userIndex = [];
      }
    }

    // 檢查是否已存在，若不存在則添加
    if (Array.isArray(userIndex) && !userIndex.includes(safeEmail)) {
      userIndex.push(safeEmail);
      cache.put(
        CACHE_KEYS.USER_INDEX,
        JSON.stringify(userIndex),
        MAX_CACHE_EXPIRATION
      );
      Logger.log("(addUserToIndex)已新增使用者到快取索引：%s", email);
    }

    return true;
  } catch (error) {
    Logger.log("(addUserToIndex)添加使用者到索引時發生錯誤：%s", error.message);
    return false;
  }
}

/**
 * @description 從快取索引中移除使用者
 * @param {string} email - 使用者電子郵件
 * @returns {boolean} 是否成功移除
 */
function removeUserFromIndex(email) {
  try {
    if (!email) {
      Logger.log("(removeUserFromIndex)無效的電子郵件");
      return false;
    }

    const cache = CacheService.getScriptCache();
    const safeEmail = getSafeKeyFromEmail(email);

    // 從快取獲取當前索引
    const indexStr = cache.get(CACHE_KEYS.USER_INDEX);
    if (!indexStr) {
      return true; // 如果索引不存在，視為移除成功
    }

    try {
      const userIndex = JSON.parse(indexStr);
      if (!Array.isArray(userIndex)) {
        Logger.log("(removeUserFromIndex)索引格式無效");
        return false;
      }

      // 從索引中移除
      const newIndex = userIndex.filter((e) => e !== safeEmail);
      cache.put(
        CACHE_KEYS.USER_INDEX,
        JSON.stringify(newIndex),
        MAX_CACHE_EXPIRATION
      );

      if (newIndex.length !== userIndex.length) {
        Logger.log("(removeUserFromIndex)已從快取索引中移除使用者：%s", email);
      }

      return true;
    } catch (e) {
      Logger.log(
        "(removeUserFromIndex)處理使用者索引時發生錯誤：%s",
        e.message
      );
      return false;
    }
  } catch (error) {
    Logger.log(
      "(removeUserFromIndex)從索引移除使用者時發生錯誤：%s",
      error.message
    );
    return false;
  }
}

/**
 * @description 取得所有使用者的快取鍵值
 * @returns {string[]} 使用者快取鍵值陣列
 */
function getAllUserCacheKeys() {
  try {
    const cache = CacheService.getScriptCache();
    const indexStr = cache.get(CACHE_KEYS.USER_INDEX);

    if (!indexStr) {
      return [];
    }

    const userIndex = JSON.parse(indexStr);
    if (!Array.isArray(userIndex)) {
      Logger.log("(getAllUserCacheKeys)索引格式無效");
      return [];
    }

    return userIndex.map(
      (safeEmail) => CACHE_KEYS.USER_DATA_PREFIX + safeEmail
    );
  } catch (error) {
    Logger.log(
      "(getAllUserCacheKeys)取得使用者快取鍵值時發生錯誤：%s",
      error.message
    );
    return [];
  }
}
