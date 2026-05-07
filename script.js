// 使用OpenWeatherMap的免费API
const API_URL = 'https://api.openweathermap.org/data/2.5';
const API_KEY = 'a8e33510e7a374318d5d6d4a5b15c3a9';

// DOM元素
const cityInput = document.getElementById('cityInput');
const searchBtn = document.getElementById('searchBtn');
const cityName = document.getElementById('cityName');
const temperature = document.querySelector('.temperature');
const weatherCondition = document.querySelector('.weather-condition');
const humidity = document.getElementById('humidity');
const wind = document.getElementById('wind');
const forecastCards = [
    document.getElementById('day1'),
    document.getElementById('day2'),
    document.getElementById('day3')
];

// 初始化应用
function init() {
    // 尝试获取用户地理位置
    getCurrentLocation();
    
    // 搜索按钮事件监听
    searchBtn.addEventListener('click', () => {
        const city = cityInput.value.trim();
        if (city) {
            getWeatherByCity(city);
        }
    });
    
    // 回车键搜索
    cityInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            const city = cityInput.value.trim();
            if (city) {
                getWeatherByCity(city);
            }
        }
    });
}

// 获取用户当前位置
function getCurrentLocation() {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(
            (position) => {
                const lat = position.coords.latitude;
                const lon = position.coords.longitude;
                getWeatherByCoords(lat, lon);
            },
            (error) => {
                console.error('获取位置失败:', error);
                // 默认获取北京天气
                getWeatherByCity('北京');
            }
        );
    } else {
        // 浏览器不支持地理位置
        getWeatherByCity('北京');
    }
}

// 根据城市名称获取天气
function getWeatherByCity(city) {
    const url = `${API_URL}/weather?q=${city}&appid=${API_KEY}&units=metric&lang=zh_cn`;
    console.log('API请求URL:', url);
    fetch(url)
        .then(response => response.json())
        .then(data => {
            console.log('API响应:', data);
            if (data.cod === 200) {
                displayWeather(data);
                getForecast(data.coord.lat, data.coord.lon);
            } else {
                console.error('城市未找到，错误代码:', data.cod, '错误信息:', data.message);
                alert(`城市未找到，请检查输入。错误信息: ${data.message}`);
            }
        })
        .catch(error => {
            console.error('获取天气失败:', error);
            alert('获取天气失败，请稍后重试');
        });
}

// 根据经纬度获取天气
function getWeatherByCoords(lat, lon) {
    const url = `${API_URL}/weather?lat=${lat}&lon=${lon}&appid=${API_KEY}&units=metric&lang=zh_cn`;
    fetch(url)
        .then(response => response.json())
        .then(data => {
            console.log('API响应:', data);
            if (data.cod === 200) {
                displayWeather(data);
                getForecast(lat, lon);
            }
        })
        .catch(error => {
            console.error('获取天气失败:', error);
        });
}

// 获取天气预报
function getForecast(lat, lon) {
    const url = `${API_URL}/forecast?lat=${lat}&lon=${lon}&appid=${API_KEY}&units=metric&lang=zh_cn&cnt=24`;
    fetch(url)
        .then(response => response.json())
        .then(data => {
            console.log('预报API响应:', data);
            if (data.cod === '200') {
                displayForecast(data.list);
            }
        })
        .catch(error => {
            console.error('获取预报失败:', error);
        });
}

// 显示当前天气
function displayWeather(data) {
    cityName.textContent = data.name;
    temperature.textContent = `${Math.round(data.main.temp)}°C`;
    weatherCondition.textContent = data.weather[0].description;
    humidity.textContent = `${data.main.humidity}%`;
    wind.textContent = `${data.wind.speed} m/s`;
}

// 显示天气预报
function displayForecast(forecastData) {
    // 每8小时一个数据点，取每天的12点左右的数据
    const dailyForecasts = [];
    for (let i = 0; i < forecastData.length; i += 8) {
        if (dailyForecasts.length < 3) {
            dailyForecasts.push(forecastData[i]);
        }
    }
    
    // 显示预报数据
    dailyForecasts.forEach((forecast, index) => {
        if (forecastCards[index]) {
            const date = new Date(forecast.dt * 1000);
            const dateStr = `${date.getMonth() + 1}/${date.getDate()}`;
            
            forecastCards[index].querySelector('.forecast-date').textContent = dateStr;
            forecastCards[index].querySelector('.forecast-temp').textContent = `${Math.round(forecast.main.temp)}°C`;
            forecastCards[index].querySelector('.forecast-condition').textContent = forecast.weather[0].description;
        }
    });
}

// 初始化应用
init();