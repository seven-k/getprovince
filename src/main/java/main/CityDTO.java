package main;

public class CityDTO {
    private String cityName;
    private String cityType;
    private String provinceName;

    public String getCityName() {
        return cityName;
    }

    public void setCityName(String cityName) {
        this.cityName = cityName;
    }

    public String getCityType() {
        return cityType;
    }

    public void setCityType(String cityType) {
        this.cityType = cityType;
    }

    public String getProvinceName() {
        return provinceName;
    }

    public void setProvinceName(String provinceName) {
        this.provinceName = provinceName;
    }

    public CityDTO(String cityName, String cityType, String provinceName) {
        this.cityName = cityName;
        this.cityType = cityType;
        this.provinceName = provinceName;
    }

    public CityDTO() {
    }
}
