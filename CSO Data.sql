CREATE TABLE `sourcedb`.`input` (
  `id` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,
  `pID` VARCHAR(20) NOT NULL,
  `ppassword` VARCHAR(15) NOT NULL,
  `playernumber` int,
  PRIMARY KEY (`id`)
)
ENGINE = InnoDB;